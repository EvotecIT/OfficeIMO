using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Removes native package encryption or selected editing restrictions.</summary>
[Cmdlet(VerbsSecurity.Unprotect, "OfficeDocument", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(WordDocument), typeof(ExcelDocument), typeof(PowerPointDocument), typeof(PdfDocument))]
public sealed class UnprotectOfficeDocumentCommand : OfficeCmdlet {
    private readonly List<(object? Document, string? Path)> _inputs = new();
    /// <summary>Input paths; supports wildcards.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path"), SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Exact input paths.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath"), Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Caller-owned native model.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public object? Document { get; set; }
    /// <summary>Current open password, or PDF owner password for removing Standard security.</summary>
    [Parameter, OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>Remove only these restrictions; omission also removes package encryption.</summary>
    [Parameter]
    public OfficeProtectionRestriction[] Restriction { get; set; } = Array.Empty<OfficeProtectionRestriction>();
    /// <summary>Write one copy; omission saves to the associated path when available.</summary>
    [Parameter]
    public string? OutputPath { get; set; }
    /// <summary>Native source policy options; PDF accepts its bounded load settings.</summary>
    [Parameter]
    public object[]? Options { get; set; }
    /// <summary>Emit the saved file or caller-owned/native replacement model.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Document") {
            var files = OfficePath.GetPipelineLiteralPaths(Document!);
            if (files != null) foreach (var path in OfficePath.ResolveInputs(this, files, literal: true)) _inputs.Add((null, path));
            else _inputs.Add((OfficeArgumentValue.Unwrap(Document) ?? throw new PSArgumentException("Document is required."), null));
        } else {
            bool literal = ParameterSetName == "LiteralPath";
            foreach (var path in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) _inputs.Add((null, path));
        }
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        int count = _inputs.Count;
        if (OutputPath != null && count != 1) throw new PSArgumentException("OutputPath requires exactly one source.");
        foreach (var restriction in Restriction) if (!Enum.IsDefined(typeof(OfficeProtectionRestriction), restriction)) throw new PSArgumentException("Invalid Restriction.");
        for (int index = 0; index < count; index++) {
            object? model = _inputs[index].Document;
            bool live = model != null;
            string? inputPath = _inputs[index].Path;
            string extension = model == null ? System.IO.Path.GetExtension(inputPath!) : OfficeProtectionDispatch.Extension(model);
            var family = OfficeProtectionDispatch.Family(extension);
            ValidateRestrictions(family);
            var read = family == OfficeProtectionFamily.Pdf ? CapturePdfReadOptions() : null;
            if (family != OfficeProtectionFamily.Pdf && Options is { Length: > 0 })
                throw new PSArgumentException("This native unprotection route has no protection-option overrides; Password authenticates the input.");
            string? output = OutputPath != null ? ResolveOutputPath(OutputPath) : model == null ? inputPath : OfficeProtectionDispatch.AssociatedPath(model);
            if (output != null && !System.IO.Path.GetExtension(output).Equals(extension, StringComparison.OrdinalIgnoreCase))
                throw new PSArgumentException("OutputPath must preserve the source format.");
            if (live && model is PdfDocument && output == null && !PassThru)
                throw new PSArgumentException("A live PDF security rewrite returns a replacement; use PassThru or OutputPath.");
            if (!ShouldProcess(output ?? "Live document", "Remove document protection")) continue;
            IDisposable? owned = null;
            try {
                model ??= owned = Load(inputPath!, family, read);
                object result = Remove(model, family, output, read, live);
                if (PassThru) WriteObject(live ? result : new FileInfo(output!));
                else if (!ReferenceEquals(result, model) && result is IDisposable replacement) replacement.Dispose();
            } finally { owned?.Dispose(); }
        }
    }

    private void ValidateRestrictions(OfficeProtectionFamily family) {
        foreach (var restriction in Restriction) if (!OfficeProtectionDispatch.Supports(family, restriction))
            throw new PSArgumentException(family + " does not support restriction " + restriction + ".");
        if (family == OfficeProtectionFamily.Pdf && Restriction.Length > 0)
            throw new PSNotSupportedException("The native PDF owner supports full security removal; selective permission removal while retaining open-password encryption requires a native operation.");
    }

    private IDisposable Load(string path, OfficeProtectionFamily family, PdfLoadOptions? pdf) => family switch {
        OfficeProtectionFamily.Word => WordDocument.Load(path, new WordLoadOptions { Password = Password.ToPlainText(), AccessMode = DocumentAccessMode.ReadWrite, PersistenceMode = DocumentPersistenceMode.Explicit, AllowLegacyProjection = false }),
        OfficeProtectionFamily.Excel => ExcelDocument.Load(path, new ExcelLoadOptions { Password = Password.ToPlainText(), AccessMode = DocumentAccessMode.ReadWrite, PersistenceMode = DocumentPersistenceMode.Explicit, AllowBinaryProjection = false }),
        OfficeProtectionFamily.PowerPoint => PowerPointDocument.Load(path, new PowerPointLoadOptions { Password = Password.ToPlainText(), AccessMode = DocumentAccessMode.ReadWrite, PersistenceMode = DocumentPersistenceMode.Explicit, AllowLegacyProjection = false }),
        _ => PdfDocument.Load(path, pdf)
    };

    private object Remove(object document, OfficeProtectionFamily family, string? output, PdfLoadOptions? read, bool live) {
        bool all = Restriction.Length == 0;
        switch (document) {
            case WordDocument word:
                CheckEditable(word.AccessMode, word.PersistenceMode, word.InspectSignatures().HasSignatures);
                if (all || word.Settings.ProtectionType is WordDocumentProtectionType type && Restriction.Any(item => item.ToString() == type.ToString())) word.Settings.RemoveProtection();
                if (output != null) {
                    var save = new WordSaveOptions(); if (all) save.Password = null;
                    PdfCommandUtilities.EnsureDirectory(output);
                    if (SamePath(word.FilePath, output)) word.Save(output, save); else word.SaveCopy(output, save);
                }
                return word;
            case ExcelDocument excel:
                CheckEditable(excel.AccessMode, excel.PersistenceMode, excel.InspectSignatures().HasSignatures);
                if (all || Restriction.Contains(OfficeProtectionRestriction.Structure) || Restriction.Contains(OfficeProtectionRestriction.Windows))
                    excel.UnprotectWorkbook(structure: all || Restriction.Contains(OfficeProtectionRestriction.Structure),
                        windows: all || Restriction.Contains(OfficeProtectionRestriction.Windows));
                if (all || Restriction.Contains(OfficeProtectionRestriction.ReadOnlyRecommended)) excel.ClearWriteReservation();
                if (output != null) {
                    var save = new ExcelSaveOptions(); if (all) save.Password = null;
                    PdfCommandUtilities.EnsureDirectory(output);
                    if (SamePath(excel.FilePath, output)) excel.Save(output, save); else excel.SaveCopy(output, save);
                }
                return excel;
            case PowerPointDocument presentation:
                CheckEditable(presentation.AccessMode, presentation.PersistenceMode, presentation.InspectSignatures().HasSignatureMetadata);
                if (output == null) throw new PSArgumentException("PowerPoint decryption requires an associated path or OutputPath.");
                PdfCommandUtilities.EnsureDirectory(output);
                if (SamePath(presentation.FilePath, output)) presentation.Save(output, new PowerPointSaveOptions { Password = null });
                else presentation.SaveCopy(output, new PowerPointSaveOptions { Password = null });
                return presentation;
            case PdfDocument pdf:
                var security = pdf.InspectGeometryForViewing().Security;
                if (security.HasSignatures) throw new PSInvalidOperationException("Native safe-save policy blocks rewriting a signed PDF.");
                if (!security.HasEncryption) {
                    if (output != null) OfficeOutput.Write(output, stream => { var bytes = pdf.ToBytes(); stream.Write(bytes, 0, bytes.Length); }, noClobber: false);
                    return pdf;
                }
                string password = Password.ToPlainText() ?? throw new PSArgumentException("Removing PDF Standard security requires the current owner Password.");
                var result = Options is { Length: > 0 } ? PdfSecurityEditor.Decrypt(pdf.ToBytes(), password, read) : pdf.Security.Decrypt(password);
                result.PreservationReport.ThrowIfFailed();
                if (output != null) OfficeOutput.Write(output, stream => stream.Write(result.Pdf, 0, result.Pdf.Length), noClobber: false);
                return live && PassThru ? result.ToDocument() : pdf;
            default: throw new PSNotSupportedException("Unsupported native Document.");
        }
    }

    private PdfLoadOptions CapturePdfReadOptions() {
        var read = new PdfLoadOptions();
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        foreach (var input in Options ?? Array.Empty<object>()) {
            var value = OfficeArgumentValue.Unwrap(input);
            if (value is Models.Pdf.PdfReadOptions shell) read = shell.Snapshot().LoadOptions;
            else if (value is PdfLoadOptions typed) read = typed.Clone();
            else if (value is System.Collections.IDictionary properties) read = (PdfLoadOptions)new OfficeOptionsTransformationAttribute(typeof(PdfLoadOptions)).ApplyTo(engine, read, properties);
            else throw new PSArgumentException("PDF unprotection accepts native bounded read options.");
        }
        return MyInvocation.BoundParameters.ContainsKey(nameof(Password)) ? read.Clone(null, Password.ToPlainText()) : read;
    }

    private static bool SamePath(string? input, string output) => input != null && OfficePath.AreEquivalent(input, output);

    private static void CheckEditable(DocumentAccessMode access, DocumentPersistenceMode persistence, bool signed) {
        if (access != DocumentAccessMode.ReadWrite) throw new PSArgumentException("Unprotection requires an editable document.");
        if (persistence != DocumentPersistenceMode.Explicit) throw new PSArgumentException("Use Explicit persistence before changing protection through this command.");
        if (signed) throw new PSInvalidOperationException("Native safe-save policy blocks rewriting signed documents.");
    }
}
