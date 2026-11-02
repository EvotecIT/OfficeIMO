using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets worksheet visibility, tab color and the active worksheet.</summary>
/// <para>Omitted properties keep their values. Sheet accepts a name, one-based position or native worksheet.
/// Inside an ExcelSheet block it defaults to the current worksheet; an explicit workbook defaults to its first sheet.</para>
/// <example><summary>Hide lookup data and open on the summary.</summary>
/// <code>Open-OfficeExcel -Path .\Pricing.xlsx {
///     Set-OfficeExcelSheet -Sheet Lookups -Visibility VeryHidden -TabColor Accent2
///     Set-OfficeExcelSheet -Sheet Summary -Active
/// }</code></example>
/// <example><summary>Clear a live worksheet's tab color.</summary>
/// <code>$book | Set-OfficeExcelSheet -Sheet 2 -NoTabColor -PassThru</code></example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelSheet", DefaultParameterSetName = "Default")]
[OutputType(typeof(ExcelSheet))]
public sealed class SetOfficeExcelSheetCommand : OfficeDocumentCmdlet<ExcelDocument> {
    /// <summary>Gets or sets a worksheet name, one-based position or native worksheet belonging to the workbook.</summary>
    [Parameter, OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Gets or sets Visible, Hidden or VeryHidden; omission preserves visibility.</summary>
    [Parameter]
    public ExcelSheetVisibility? Visibility { get; set; }

    /// <summary>Gets or sets the literal or theme-based tab color.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeColorTransformation]
    public OfficeColor? TabColor { get; set; }

    /// <summary>Gets or sets whether to remove the tab color; exclusive with TabColor.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "NoTabColor")]
    public SwitchParameter NoTabColor { get; set; }

    /// <summary>Gets or sets whether this visible worksheet opens as the workbook's active worksheet.</summary>
    [Parameter]
    public SwitchParameter Active { get; set; }

    /// <summary>Gets or sets whether to emit the edited native worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            if (Visibility.HasValue && !Enum.IsDefined(typeof(ExcelSheetVisibility), Visibility.Value))
                throw new ArgumentOutOfRangeException(nameof(Visibility));
            if (Active && Visibility.HasValue && Visibility != ExcelSheetVisibility.Visible)
                throw new ArgumentException("A hidden worksheet cannot be the active worksheet.");
            ExcelSheet sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
            if (Active && sheet.Hidden && Visibility != ExcelSheetVisibility.Visible)
                throw new ArgumentException("Show the worksheet before making it active, or specify -Visibility Visible.");

            if (Visibility == ExcelSheetVisibility.Visible) sheet.SetHidden(false);
            else if (Visibility == ExcelSheetVisibility.Hidden) sheet.SetHidden(true);
            else if (Visibility == ExcelSheetVisibility.VeryHidden) sheet.SetVeryHidden(true);
            if (Active) sheet.Document.SetActiveWorksheet(sheet);
            if (TabColor.HasValue) sheet.SetTabColor(TabColor.Value);
            else if (NoTabColor) sheet.ClearTabColor();
            if (PassThru) WriteObject(sheet);
        } catch (ArgumentException error) { InvalidArgument(error, Document); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "excel-sheet-edit-failed", Document);
        }
    }
}
