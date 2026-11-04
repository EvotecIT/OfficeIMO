using System;
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
[Cmdlet(VerbsCommon.Move, "OfficeExcelSheet", DefaultParameterSetName = "Default", SupportsShouldProcess = true)]
[OutputType(typeof(ExcelSheet))]
public sealed class MoveOfficeExcelSheetCommand : OfficeDocumentCmdlet<ExcelDocument> {
    /// <summary>Gets or sets a worksheet name, one-based position or native worksheet.</summary>
    [Parameter, OfficeSheetTransformation, ValidateNotNull]
    public OfficeSheetSelector? Sheet { get; set; }
    /// <summary>Gets or sets the one-based destination position.</summary>
    [Parameter(Mandatory = true), ValidateRange(1, int.MaxValue)]
    public int Position { get; set; }
    /// <summary>Gets or sets whether to return the moved native worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            ExcelSheet sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
            if (Position > sheet.Document.Sheets.Count) throw new ArgumentOutOfRangeException(nameof(Position), "Position is outside the workbook.");
            if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, sheet.Document, null, "Move worksheet " + sheet.Name)) return;
            sheet.Document.ReorderWorksheet(sheet, Position - 1);
            if (PassThru) WriteObject(sheet);
        } catch (ArgumentException error) { InvalidArgument(error, Document); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "excel-sheet-move-failed", Document);
        }
    }
}
