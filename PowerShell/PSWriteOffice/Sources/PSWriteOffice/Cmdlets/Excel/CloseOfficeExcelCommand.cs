using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Closes a live Excel workbook without saving pending changes.</summary>
/// <para>Use Save-OfficeExcel before closing to retain changes. Workbooks with SaveOnDispose are rejected so disposal cannot publish a file.</para>
/// <example><summary>Save edits explicitly, then close.</summary>
/// <code>$workbook | Save-OfficeExcel
/// $workbook | Close-OfficeExcel</code></example>
/// <example>
/// <summary>Close a workbook without saving</summary>
/// <code>$book = Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx -ReadOnly
/// Close-OfficeExcel -Document $book</code>
/// <para>Releases a workbook that was only read; nothing is written to disk.</para>
/// </example>
/// <example>
/// <summary>Always close after reading</summary>
/// <code>$book = Open-OfficeExcel -Path .\Finance\Purchase-Requests.xlsx -ReadOnly
/// try {
///     Import-OfficeExcel -Document $book -Table Requests | Where-Object Status -eq 'Awaiting approval'
/// } finally {
///     Close-OfficeExcel -Document $book
/// }</code>
/// <para>Reads the open purchase requests and closes the workbook in a finally block, so the file is released even if the filter fails.</para>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficeExcel", DefaultParameterSetName = "Default")]
public sealed class CloseOfficeExcelCommand : PSCmdlet {
    /// <summary>Explicit-persistence workbook to close.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [ValidateNotNull]
    public ExcelDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        ExcelDocumentService.EnsureExplicitPersistence(Document);
        ExcelDocumentService.CloseDocument(Document);
        OfficeUnsavedNewDocuments.WarnIfUnsaved(this, Document, "workbook", "Excel");
    }
}
