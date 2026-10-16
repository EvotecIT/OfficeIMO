using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes values or applies column layout to a worksheet in a DSL block or through the pipeline.</summary>
/// <example>
///   <summary>Populate a column and auto-fit it.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelColumn -Column 1 -Values 'North','South' -AutoFit }</code>
///   <para>Writes values into column A and adjusts the width.</para>
/// </example>
/// <example>
///   <summary>Set a readable report column width through the worksheet pipeline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sheet | Set-OfficeExcelColumn -ColumnName C -Width 48</code>
///   <para>Sets column C to 48 Excel character units. Save the workbook after editing.</para>
/// </example>
/// <example>
///   <summary>Unhide a column in a named worksheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook | Set-OfficeExcelColumn -Sheet 'Review' -ColumnName C -Hidden $false -AutoFit</code>
///   <para>Uses the workbook pipeline target without an active DSL scope.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelColumn")]
[Alias("ExcelColumn")]
public sealed class SetOfficeExcelColumnCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Worksheet to modify outside a DSL context.</summary>
    [Parameter(ValueFromPipeline = true)]
    [Alias("SheetObject")]
    public ExcelSheet? Worksheet { get; set; }

    /// <summary>Workbook to modify. Accepts pipeline input; select its worksheet with Sheet or SheetIndex.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name when using Document. The first worksheet is used when no selector is supplied.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Zero-based worksheet index when using Document.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter(Position = 0)]
    public int? Column { get; set; }

    /// <summary>Column letter reference (e.g., A, BC).</summary>
    [Parameter]
    [Alias("ColumnLetter", "Letter")]
    public string? ColumnName { get; set; }

    /// <summary>Values to write down the column. Numeric, Boolean, and date values retain their native Excel types; strings remain text.</summary>
    [Parameter]
    public object[]? Values { get; set; }

    /// <summary>Starting row index (1-based) for values.</summary>
    [Parameter]
    public int StartRow { get; set; } = 1;

    /// <summary>Column width in Excel character units.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Hide or show the column.</summary>
    [Parameter]
    public bool? Hidden { get; set; }

    /// <summary>Auto-fit the column width after updates.</summary>
    [Parameter]
    public SwitchParameter AutoFit { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Worksheet, Document, Sheet, SheetIndex);

        var columnIndex = ExcelHostExtensions.ResolveColumnIndex(Column, ColumnName);
        if (columnIndex < 1) {
            throw new ArgumentOutOfRangeException(nameof(Column), "Column index must be 1 or greater.");
        }

        var hasAction = false;

        if (Values != null && Values.Length > 0) {
            if (StartRow < 1) {
                throw new ArgumentOutOfRangeException(nameof(StartRow), "StartRow must be 1 or greater.");
            }

            var cells = new List<(int Row, int Column, object Value)>(Values.Length);
            for (int i = 0; i < Values.Length; i++) {
                var value = PowerShellObjectNormalizer.NormalizeCellValueForTable(Values[i]) ?? string.Empty;
                cells.Add((StartRow + i, columnIndex, value));
            }
            sheet.CellValues(cells);
            hasAction = true;
        }

        if (Width.HasValue) {
            sheet.SetColumnWidth(columnIndex, Width.Value);
            hasAction = true;
        }

        if (Hidden.HasValue) {
            sheet.SetColumnHidden(columnIndex, Hidden.Value);
            hasAction = true;
        }

        if (AutoFit.IsPresent) {
            sheet.AutoFitColumn(columnIndex);
            hasAction = true;
        }

        if (!hasAction) {
            throw new PSArgumentException("Provide -Values, -Width, -Hidden, or -AutoFit to update the column.");
        }
        WritePassThru(sheet);
    }
}
