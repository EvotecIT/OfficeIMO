using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Foundation;

/// <summary>Reads worksheet snapshots through shared literal/wildcard admission and caller-owned document scopes.</summary>
public abstract class ExcelQueryCmdlet : OfficeCmdlet {
    /// <summary>Workbook paths, including filesystem wildcard patterns.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal workbook paths. FileInfo.FullName and provider PSPath bind here.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet names, wildcard patterns, one-based positions or native worksheet objects.</summary>
    [Parameter, OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <inheritdoc />
    protected sealed override void ProcessRecord() {
        if (ParameterSetName == "Document") {
            ExcelDocument document = Document ?? OfficeDocumentContext.Find<ExcelDocument>()
                ?? ExcelDslContext.Require(this).Document;
            WriteResults(document, Document == null ? ExcelDslContext.Current?.CurrentSheet : null, null);
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string path in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) {
            ExcelDocument document;
            try {
                document = ExcelDocument.Load(path, new ExcelLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    PersistenceMode = DocumentPersistenceMode.Explicit
                });
            } catch (Exception error) when (error is not PipelineStoppedException and not ActionPreferenceStopException
                and not OperationCanceledException and not OutOfMemoryException and not StackOverflowException) {
                ItemError(error, "excel-query-failed", path);
                continue;
            }
            using (document) {
                try {
                    WriteResults(document, null, path);
                } catch (Exception error) when (error is not ArgumentException and not WildcardPatternException and not PipelineStoppedException
                    and not ActionPreferenceStopException and not OperationCanceledException
                    and not OutOfMemoryException and not StackOverflowException) {
                    ItemError(error, "excel-query-failed", path);
                }
            }
        }
    }

    private void WriteResults(ExcelDocument document, ExcelSheet? defaultSheet, string? path) {
        try {
            foreach (PSObject record in Query(document, ExcelSheetResolver.ResolveMany(document, Sheet, defaultSheet), path))
                WriteObject(record);
        } catch (Exception error) when (error is ArgumentException or WildcardPatternException) {
            InvalidArgument(error, (object?) path ?? document);
        }
    }

    /// <summary>Projects the native snapshots for selected worksheets; no workbook ownership is transferred.</summary>
    protected abstract IEnumerable<PSObject> Query(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets, string? path);

    /// <summary>Keeps the native snapshot as BaseObject while adding its worksheet and file provenance for the pipeline.</summary>
    protected static PSObject CreateRecord(object value, string? path, string? sheetName = null) {
        PSObject result = PSObject.AsPSObject(value);
        if (path != null) result.Properties.Add(new PSNoteProperty("Path", path));
        if (sheetName != null) {
            result.Properties.Add(new PSNoteProperty("SheetName", sheetName));
            result.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        }
        return result;
    }
}
