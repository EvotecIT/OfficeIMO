using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Moves a worksheet to a one-based workbook position.</summary>
/// <para>Uses the explicit or enclosing live workbook. Save it explicitly to persist the new order.</para>
/// <example><summary>Move the summary worksheet to the front.</summary>
/// <code>$book | Move-OfficeExcelSheet -Sheet Summary -Position 1
/// Save-OfficeExcel -Document $book</code></example>
/// <example>
/// <summary>Move a sheet to the front</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     Move-OfficeExcelSheet -Sheet Summary -Position 1
/// }</code>
/// <para>Moves the Summary sheet to the first tab position.</para>
/// </example>
/// <example>
/// <summary>Order monthly sheets newest first</summary>
/// <code>$book = Open-OfficeExcel -Path .\Reports\Tickets-2026.xlsx
/// foreach ($month in 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October') {
///     Move-OfficeExcelSheet -Document $book -Sheet $month -Position 1
/// }
/// Save-OfficeExcel -Document $book -PassThru | Close-OfficeExcel</code>
/// <para>Moves the sheet for each month to the front in turn, so the most recent month ends up first.</para>
/// </example>
[Cmdlet(VerbsCommon.Move, "OfficeExcelSheet", DefaultParameterSetName = "Default")]
[OutputType(typeof(ExcelSheet))]
public sealed class MoveOfficeExcelSheetCommand : OfficeDocumentCmdlet<ExcelDocument> {
    /// <summary>Gets or sets a worksheet name, one-based position or native worksheet.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation, ValidateNotNull]
    [Alias("WorksheetName")]
    public OfficeSheetSelector? Sheet { get; set; }
    /// <summary>Gets or sets the one-based destination position.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    [ValidateRange(1, int.MaxValue)]
    public int Position { get; set; }
    /// <summary>Gets or sets whether to return the moved native worksheet.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            ExcelSheet sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
            if (Position > sheet.Document.Sheets.Count) throw new ArgumentOutOfRangeException(nameof(Position), "Position is outside the workbook.");
            sheet.Document.ReorderWorksheet(sheet, Position - 1);
            if (PassThru) WriteObject(sheet);
        } catch (ArgumentException error) { InvalidArgument(error, Document); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "excel-sheet-move-failed", Document);
        }
    }
}
