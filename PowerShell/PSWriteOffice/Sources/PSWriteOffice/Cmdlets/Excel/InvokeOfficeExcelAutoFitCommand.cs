using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Fits column widths and row heights to their content.</summary>
/// <example>
/// <summary>Fit every column and row</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
///     ExcelSheet 'Servers' {
///         ExcelAutoFit
///     }
/// }</code>
/// <para>Fits all columns and rows of the current sheet to their content.</para>
/// </example>
/// <example>
/// <summary>Fit only the columns you changed</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Changes.xlsx {
///     ExcelSheet 'Changes' {
///         ExcelColumn F -Values (Import-OfficeCsv -Path .\Data\Change-Notes.csv).Note -StartRow 2
///         ExcelAutoFit -Column E, F
///     }
/// }</code>
/// <para>Writes long notes into column F and refits only columns E and F, leaving the widths of the rest as designed.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeExcelAutoFit", DefaultParameterSetName = "Default")]
[Alias("ExcelAutoFit")]
public sealed class InvokeOfficeExcelAutoFitCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Auto-fit specific column indexes (1-based).</summary>
    [Parameter(ParameterSetName = "Default")]
    public object[]? Column { get; set; }

    /// <summary>Auto-fit specific row indexes (1-based).</summary>
    [Parameter(ParameterSetName = "Default")]
    public object[]? Row { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        bool all = Column == null && Row == null;
        if (all || Column?.Any(value => Convert.ToString(value) == "*") == true) sheet.AutoFitColumns();
        else if (Column != null) sheet.AutoFitColumnsFor(Column.Select(ExcelCommandTargets.Column));
        if (all || Row?.Any(value => Convert.ToString(value) == "*") == true) sheet.AutoFitRows();
        else if (Row != null) foreach (object value in Row) sheet.AutoFitRow(Convert.ToInt32(value, CultureInfo.InvariantCulture));
    }

}
