using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Closes a live Excel workbook without saving pending changes.</summary>
/// <para>Use Save-OfficeExcel before closing to retain changes. Workbooks with SaveOnDispose are rejected so disposal cannot publish a file.</para>
/// <example><summary>Save edits explicitly, then close.</summary>
/// <code>$workbook | Save-OfficeExcel
/// $workbook | Close-OfficeExcel</code></example>
[Cmdlet(VerbsCommon.Close, "OfficeExcel", DefaultParameterSetName = "Default")]
public sealed class CloseOfficeExcelCommand : PSCmdlet {
    /// <summary>Explicit-persistence workbook to close.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true), ValidateNotNull]
    public ExcelDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        ExcelDocumentService.EnsureExplicitPersistence(Document);
        ExcelDocumentService.CloseDocument(Document);
    }
}
