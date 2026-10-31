using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes values and applies row layout to a worksheet in a DSL block or through the pipeline.</summary>
/// <example>
///   <summary>Write a row starting at column A.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelRow -Row 2 -Values 'North', 1200 }</code>
///   <para>Writes two values into row 2, columns A and B.</para>
/// </example>
/// <example>
///   <summary>Wrap a report row and adjust its height through the worksheet pipeline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sheet | Set-OfficeExcelRow -Row 2 -WrapText $true -AutoFit</code>
///   <para>Preserves existing cell formatting while wrapping the used row span and fitting its height. Save the workbook after editing.</para>
/// </example>
/// <example>
///   <summary>Update a named worksheet through the workbook pipeline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook | Set-OfficeExcelRow -Sheet 'Review' -Row 2 -Values 'Identity', 'Ready'</code>
///   <para>Selects the worksheet by name without opening a DSL scope.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelRow")]
[Alias("ExcelRow")]
public sealed class SetOfficeExcelRowCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
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

    /// <summary>1-based row index.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public int Row { get; set; }

    /// <summary>Values to write across the row. Numeric, Boolean, and date values retain their native Excel types; strings remain text.</summary>
    [Parameter(Position = 1)]
    public object[]? Values { get; set; }

    /// <summary>Starting column index (1-based).</summary>
    [Parameter]
    public int StartColumn { get; set; } = 1;

    /// <summary>Explicit row height in points.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Clear any custom row height.</summary>
    [Parameter]
    public SwitchParameter ClearHeight { get; set; }

    /// <summary>Auto-fit row height after writing values and applying style.</summary>
    [Parameter]
    public SwitchParameter AutoFit { get; set; }

    /// <summary>Hide or unhide the row. Use -Hidden:$false to unhide.</summary>
    [Parameter]
    public bool? Hidden { get; set; }

    /// <summary>Apply or clear bold styling across the target row span.</summary>
    [Parameter]
    public bool? Bold { get; set; }

    /// <summary>Apply or clear italic styling across the target row span.</summary>
    [Parameter]
    public bool? Italic { get; set; }

    /// <summary>Apply or clear underline styling across the target row span.</summary>
    [Parameter]
    public bool? Underline { get; set; }

    /// <summary>Apply or clear wrap text across the target row span.</summary>
    [Parameter]
    public bool? WrapText { get; set; }

    /// <summary>Apply a font family across the target row span.</summary>
    [Parameter]
    public string? FontName { get; set; }

    /// <summary>Apply a background color across the target row span.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }

    /// <summary>First 1-based column affected by style options.</summary>
    [Parameter]
    public int? FirstColumn { get; set; }

    /// <summary>Last 1-based column affected by style options.</summary>
    [Parameter]
    public int? LastColumn { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Worksheet, Document, Sheet, SheetIndex);

        if (Row < 1) {
            throw new ArgumentOutOfRangeException(nameof(Row), "Row index must be 1 or greater.");
        }

        if (StartColumn < 1) {
            throw new ArgumentOutOfRangeException(nameof(StartColumn), "StartColumn must be 1 or greater.");
        }

        var values = Values ?? Array.Empty<object>();
        var hasLayout = HasLayoutOptions();
        if (values.Length == 0 && !hasLayout) {
            throw new PSArgumentException("Provide row values or at least one layout/style option.", nameof(Values));
        }

        if (values.Length > 0) {
            var cells = new List<(int Row, int Column, object Value)>(values.Length);
            for (int i = 0; i < values.Length; i++) {
                var value = PowerShellObjectNormalizer.NormalizeCellValueForTable(values[i]) ?? string.Empty;
                cells.Add((Row, StartColumn + i, value));
            }

            sheet.CellValues(cells);
        }

        if (hasLayout) {
            sheet.SetRowLayout(Row, new ExcelRowLayoutOptions {
                Height = Height.HasValue ? ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(Height.Value)) : (ExcelColumnWidth?)null,
                ClearHeight = ClearHeight.IsPresent,
                AutoFit = AutoFit.IsPresent,
                Hidden = Hidden,
                Bold = Bold,
                Italic = Italic,
                Underline = Underline,
                WrapText = WrapText,
                FontName = FontName,
                BackgroundColor = BackgroundColor,
                FirstColumn = FirstColumn ?? (values.Length > 0 ? StartColumn : null),
                LastColumn = LastColumn ?? (values.Length > 0 ? StartColumn + values.Length - 1 : null)
            });
        }
        WritePassThru(sheet);
    }

    private bool HasLayoutOptions() {
        return Height.HasValue ||
            ClearHeight.IsPresent ||
            AutoFit.IsPresent ||
            Hidden.HasValue ||
            Bold.HasValue ||
            Italic.HasValue ||
            Underline.HasValue ||
            WrapText.HasValue ||
            !string.IsNullOrWhiteSpace(FontName) ||
            !string.IsNullOrWhiteSpace(BackgroundColor) ||
            FirstColumn.HasValue ||
            LastColumn.HasValue;
    }
}
