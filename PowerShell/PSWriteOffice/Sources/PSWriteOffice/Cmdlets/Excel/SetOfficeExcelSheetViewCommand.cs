using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets frozen panes, gridlines, direction, zoom and view mode for a worksheet.</summary>
/// <para>Omitted settings keep their values; explicitly false switches reverse their setting.
/// Set both frozen dimensions to zero to unfreeze the sheet.</para>
/// <example><summary>Keep report headers and labels visible.</summary>
/// <code>ExcelSheet 'Dashboard' { ExcelSheetView -FreezeRows 3 -FreezeColumns 1 -NoGridlines -Zoom 85 }</code></example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelSheetView", DefaultParameterSetName = "Default")]
[Alias("ExcelSheetView")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelSheetViewCommand : OfficeDocumentCmdlet<ExcelDocument> {
    /// <summary>Gets or sets a worksheet name, one-based position or native worksheet.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }
    /// <summary>Gets or sets the number of frozen top rows; omission preserves the current row count.</summary>
    [Parameter, ValidateRange(0, 1048575)]
    public int? FreezeRows { get; set; }
    /// <summary>Gets or sets the number of frozen left columns; omission preserves the current column count.</summary>
    [Parameter, ValidateRange(0, 16383)]
    public int? FreezeColumns { get; set; }
    /// <summary>Gets or sets whether gridlines are hidden; explicitly false shows them.</summary>
    [Parameter]
    public SwitchParameter NoGridlines { get; set; }
    /// <summary>Gets or sets right-to-left layout; explicitly false restores left-to-right.</summary>
    [Parameter]
    public SwitchParameter RightToLeft { get; set; }
    /// <summary>Gets or sets the current zoom percentage from 10 to 400.</summary>
    [Parameter, ValidateRange(10, 400)]
    public uint? Zoom { get; set; }
    /// <summary>Gets or sets the normal-view zoom percentage from 10 to 400.</summary>
    [Parameter, ValidateRange(10, 400)]
    public uint? ZoomNormal { get; set; }
    /// <summary>Gets or sets the worksheet view mode.</summary>
    [Parameter]
    public ExcelWorksheetViewKind? View { get; set; }
    /// <summary>Gets or sets whether to emit the edited native worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            if (View.HasValue && !Enum.IsDefined(typeof(ExcelWorksheetViewKind), View.Value))
                throw new ArgumentOutOfRangeException(nameof(View));
            ExcelSheet sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
            bool? hidden = BoundSwitch(nameof(NoGridlines));
            bool? rightToLeft = BoundSwitch(nameof(RightToLeft));
            if (hidden.HasValue || rightToLeft.HasValue || Zoom.HasValue || ZoomNormal.HasValue || View.HasValue)
                sheet.SetViewOptions(showGridlines: hidden.HasValue ? !hidden.Value : null,
                    rightToLeft: rightToLeft, zoomScale: Zoom, zoomScaleNormal: ZoomNormal, view: View);
            if (FreezeRows.HasValue || FreezeColumns.HasValue) {
                ExcelWorksheetViewInfo current = sheet.GetViewInfo();
                sheet.Freeze(FreezeRows ?? current.FrozenRowCount, FreezeColumns ?? current.FrozenColumnCount);
            }
            if (PassThru) WriteObject(sheet);
        } catch (ArgumentException error) { InvalidArgument(error, Document); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "excel-sheet-view-edit-failed", Document);
        }
    }
}
