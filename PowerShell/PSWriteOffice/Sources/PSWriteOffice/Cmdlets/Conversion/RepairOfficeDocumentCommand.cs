using System.Collections.Generic;
using System;
using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Runs native workbook repairs in place, into a copy, or on a caller-owned model.</summary>
[Cmdlet(VerbsDiagnostic.Repair, "OfficeDocument", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeRepairReport))]
public sealed class RepairOfficeDocumentCommand : OfficeCmdlet {
    private readonly List<(object? Document, string? Path)> _inputs = new();
    /// <summary>Input paths; supports wildcards.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path"), SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Exact input paths.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath"), Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Caller-owned native document to repair without taking ownership.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public object? Document { get; set; }
    /// <summary>Write one repaired copy; file inputs otherwise save in place.</summary>
    [Parameter]
    public string? OutputPath { get; set; }
    /// <summary>Native format repair settings and property overrides.</summary>
    [Parameter]
    public object[]? Options { get; set; }
    /// <summary>Password used to open an encrypted file.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath"), OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>Emit native repair findings and publication details.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Document") {
            var files = OfficePath.GetPipelineLiteralPaths(Document!);
            if (files != null) foreach (var path in OfficePath.ResolveInputs(this, files, literal: true)) _inputs.Add((null, path));
            else _inputs.Add((OfficeArgumentValue.Unwrap(Document), null));
        } else {
            bool literal = ParameterSetName == "LiteralPath";
            foreach (var path in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) _inputs.Add((null, path));
        }
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        var options = CaptureOptions();
        if (OutputPath != null && _inputs.Count != 1) throw new PSArgumentException("OutputPath requires one input file or model.");
        foreach (var input in _inputs) {
            if (input.Document != null) {
                if (input.Document is not ExcelDocument live) throw new PSNotSupportedException("Native automatic repair currently supports Excel workbooks.");
                Repair(live, null, OutputPath == null ? null : ResolveOutputPath(OutputPath), options);
                continue;
            }
            string path = input.Path!;
            if (!ExcelFormatCatalog.TryGetByExtension(System.IO.Path.GetExtension(path), out _)) throw new PSNotSupportedException("Native automatic repair currently supports Excel workbooks.");
            string output = OutputPath == null ? path : ResolveOutputPath(OutputPath);
            if (!ShouldProcess(output, "Repair document and save")) continue;
            using var workbook = ExcelDocument.Load(path, new ExcelLoadOptions {
                AccessMode = DocumentAccessMode.ReadWrite, PersistenceMode = DocumentPersistenceMode.Explicit,
                AllowBinaryProjection = false, Password = Password.ToPlainText()
            });
            Repair(workbook, path, output, options, approved: true);
        }
    }

    private void Repair(ExcelDocument workbook, string? input, string? output, ExcelWorkbookRepairOptions options, bool approved = false) {
        if (workbook.AccessMode != DocumentAccessMode.ReadWrite) throw new PSArgumentException("Repair requires an editable workbook.");
        if (workbook.InspectSignatures().HasSignatures) throw new PSInvalidOperationException("Native safe-save policy blocks repair of signed workbooks.");
        if (!approved && !ShouldProcess(output ?? workbook.FilePath ?? "Live workbook", "Repair document")) return;
        var report = workbook.RepairWorkbook(options);
        if (output != null) {
            using var bytes = new MemoryStream();
            workbook.Save(bytes, new ExcelSaveOptions());
            bytes.Position = 0;
            OfficeOutput.Write(output, target => bytes.CopyTo(target), noClobber: false);
        }
        if (PassThru) WriteObject(new OfficeRepairReport(report, input, output));
    }

    private ExcelWorkbookRepairOptions CaptureOptions() {
        var result = new ExcelWorkbookRepairOptions();
        var binder = new OfficeOptionsTransformationAttribute(typeof(ExcelWorkbookRepairOptions));
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        foreach (var input in Options ?? Array.Empty<object>()) {
            object? value = OfficeArgumentValue.Unwrap(input);
            if (value is ExcelWorkbookRepairOptions typed) result = Copy(typed);
            else if (value is System.Collections.IDictionary table) result = (ExcelWorkbookRepairOptions)binder.ApplyTo(engine, result, table);
            else throw new PSArgumentException("Options must be native ExcelWorkbookRepairOptions or a property table.");
        }
        result.Save = false; // Publication belongs to this command's single native ShouldProcess decision.
        return result;
    }

    private static ExcelWorkbookRepairOptions Copy(ExcelWorkbookRepairOptions source) => new() {
        DefinedNames = source.DefinedNames,
        Tables = source.Tables,
        SheetViews = source.SheetViews,
        PrintSettings = source.PrintSettings,
        Protection = source.Protection,
        AutoFilters = source.AutoFilters,
        Drawings = source.Drawings,
        LinksAndComments = source.LinksAndComments,
        Calculation = source.Calculation,
        WorkbookArtifacts = source.WorkbookArtifacts,
        Save = false
    };
}
