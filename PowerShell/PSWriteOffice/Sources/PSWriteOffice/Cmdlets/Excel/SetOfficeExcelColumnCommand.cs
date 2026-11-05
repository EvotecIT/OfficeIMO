using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes values down a column and sets its width, visibility, number format and style, by position or header.</summary>
/// <example>
/// <summary>Set a column width</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Changes.xlsx {
///     ExcelSheet 'Changes' {
///         ExcelColumn B -Width 60 -WrapText
///     }
/// }</code>
/// <para>Makes column B wide enough for long descriptions and wraps nothing else.</para>
/// </example>
/// <example>
/// <summary>Style table columns by header</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelSheet 'Orders' {
///         ExcelColumn -Header Amount -Table Orders -Style Currency -Decimals 2
///         ExcelColumn -Header Status -Table Orders -BackgroundByText @{ Overdue = '#FEE2E2'; Paid = '#DCFCE7' }
///     }
/// }</code>
/// <para>Formats the Amount column of the Orders table as currency and colours the Status column cells by their text.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelColumn", DefaultParameterSetName = "Column")]
[Alias("ExcelColumn")]
public sealed class SetOfficeExcelColumnCommand : PSCmdlet {
    /// <summary>1-based column index.</summary>
    [Parameter(ParameterSetName = "Column", Mandatory = true, Position = 0)]
    public object Column { get; set; } = null!;

    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Values to write down the column. Numeric, Boolean, and date values retain their native Excel types; strings remain text.</summary>
    [Parameter()]
    public object[]? Values { get; set; }

    /// <summary>Starting row index (1-based) for values.</summary>
    [Parameter()]
    [ValidateRange(1, 1048576)]
    public int StartRow { get; set; } = 1;

    /// <summary>Column width in Excel character units.</summary>
    [Parameter()]
    [ExcelColumnWidthTransformation]
    public ExcelColumnWidth? Width { get; set; }

    /// <summary>Hide or show the column.</summary>
    [Parameter()]
    public SwitchParameter Hidden { get; set; }

    /// <summary>Auto-fit the column width after updates.</summary>
    [Parameter()]
    public SwitchParameter AutoFit { get; set; }

    /// <summary>Custom number format. Also used when <see cref="Style"/> is NumberFormat.</summary>
    [Parameter()]
    public string? NumberFormat { get; set; }

    /// <summary>Preset number style to apply.</summary>
    [Parameter()]
    public OfficeExcelColumnStylePreset? Style { get; set; }

    /// <summary>Decimal places for number, percent, and currency styles.</summary>
    [Parameter()]
    [ValidateRange(0, 30)]
    public int Decimals { get; set; } = 2;

    /// <summary>Culture used for value conversion and number-format presets.</summary>
    [Parameter()]
    public CultureInfo? Culture { get; set; }

    /// <summary>Date or DateTime number format pattern.</summary>
    [Parameter()]
    public string? Pattern { get; set; }

    /// <summary>Applies bold text. An explicit false value clears bold; omission preserves the native value.</summary>
    [Parameter()]
    public SwitchParameter Bold { get; set; }

    /// <summary>Cell fill colour expressed as a literal OfficeColor or theme role.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }

    /// <summary>Text or rule colour expressed as a literal OfficeColor or theme role.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <summary>Align cell content in the resolved column.</summary>
    [Parameter()]
    public OfficeExcelColumnAlignment? Alignment { get; set; }

    /// <summary>Apply or clear wrap text across the target row span.</summary>
    [Parameter()]
    public SwitchParameter WrapText { get; set; }

    /// <summary>Background colors keyed by matching cell text.</summary>
    [Parameter()]
    public Hashtable? BackgroundByText { get; set; }

    /// <summary>Maps matching cell text to font colours or theme roles.</summary>
    [Parameter()]
    public Hashtable? ColorByText { get; set; }

    /// <summary>Values that should be bolded when the cell text matches.</summary>
    [Parameter()]
    public string[]? BoldByText { get; set; }

    /// <summary>Matches names and text using ordinal case sensitivity. Omission ignores case.</summary>
    [Parameter()]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Header text identifying the target column.</summary>
    [Parameter(ParameterSetName = "Header", Mandatory = true)]
    public string Header { get; set; } = null!;

    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "Header")]
    public string? Table { get; set; }

    /// <summary>Includes the header cell in the resolved column range.</summary>
    [Parameter(ParameterSetName = "Header")]
    public SwitchParameter IncludeHeader { get; set; }

    /// <summary>Do nothing when the header cannot be found instead of throwing.</summary>
    [Parameter(ParameterSetName = "Header")]
    public SwitchParameter IgnoreMissing { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        int column;
        string target;
        if (ParameterSetName == "Header") {
            if (!sheet.TryGetColumnRangeByHeader(Header!, Table, 0, IncludeHeader, true, out target)) {
                if (IgnoreMissing) return;
                throw new PSArgumentException($"Header '{Header}' does not exist on worksheet '{sheet.Name}'.");
            }
            column = ExcelA1.ParseRange(target).Item2;
        } else {
            column = ExcelCommandTargets.Column(Column);
            var (_, _, last, _) = ExcelA1.ParseRange(sheet.UsedRange.Address);
            target = ExcelA1.CellReference(1, column) + ":" + ExcelA1.CellReference(Math.Max(Math.Max(1, last), Values == null ? 1 : checked(StartRow + Values.Length - 1)), column);
        }
        if (Values != null) for (int index = 0; index < Values.Length; index++) sheet.Cell(StartRow + index, column).SetValue(PowerShellObjectNormalizer.NormalizeCellValueForTable(Values[index]) ?? string.Empty);
        if (Width.HasValue) sheet.SetColumnWidth(column, Width.Value);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Hidden))) sheet.SetColumnHidden(column, Hidden);
        var range = sheet.Range(target);
        if (NumberFormat != null) range.SetNumberFormat(NumberFormat);
        if (Style.HasValue) {
            string format = Style.Value switch {
                OfficeExcelColumnStylePreset.Integer => ExcelNumberFormats.Get(ExcelNumberPreset.Integer), OfficeExcelColumnStylePreset.Text => ExcelNumberFormats.Get(ExcelNumberPreset.Text), OfficeExcelColumnStylePreset.Percent => ExcelNumberFormats.Get(ExcelNumberPreset.Percent, Decimals),
                OfficeExcelColumnStylePreset.Currency => ExcelNumberFormats.Get(ExcelNumberPreset.Currency, Decimals, Culture ?? CultureInfo.CurrentCulture),
                OfficeExcelColumnStylePreset.Date => Pattern ?? "yyyy-mm-dd", OfficeExcelColumnStylePreset.DateTime => Pattern ?? "yyyy-mm-dd hh:mm:ss",
                OfficeExcelColumnStylePreset.Time => ExcelNumberFormats.Get(ExcelNumberPreset.Time), OfficeExcelColumnStylePreset.DurationHours => ExcelNumberFormats.Get(ExcelNumberPreset.DurationHours),
                OfficeExcelColumnStylePreset.NumberFormat => NumberFormat ?? throw new PSArgumentException("NumberFormat is required for the NumberFormat preset."),
                _ => ExcelNumberFormats.Get(ExcelNumberPreset.Decimal, Decimals)
            };
            range.SetNumberFormat(format);
        }
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Bold))) range.SetBold(Bold);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(WrapText))) range.SetWrapText(WrapText);
        if (BackgroundColor.HasValue) range.SetFillColor(BackgroundColor.Value);
        if (Color.HasValue) range.SetColor(Color.Value);
        if (Alignment.HasValue) range.SetHorizontalAlignment(Alignment == OfficeExcelColumnAlignment.Left ? ExcelHorizontalAlignment.Left : Alignment == OfficeExcelColumnAlignment.Right ? ExcelHorizontalAlignment.Right : ExcelHorizontalAlignment.Center);
        if (BackgroundByText != null || ColorByText != null || BoldByText != null) {
            var (first, _, last, _) = ExcelA1.ParseRange(target);
            for (int row = first; row <= last; row++) {
                var cell = sheet.Cell(row, column);
                string text = cell.GetFormattedText();
                ApplyColorMap(BackgroundByText, text, color => cell.SetFillColor(color));
                ApplyColorMap(ColorByText, text, color => cell.SetColor(color));
                if (BoldByText?.Contains(text, CaseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase) == true) cell.SetBold();
            }
        }
        if (AutoFit) sheet.AutoFitColumn(column);
        if (PassThru) WriteObject(sheet);
    }
    private void ApplyColorMap(Hashtable? map, string text, Action<OfficeColor> apply) {
        if (map == null) return;
        foreach (DictionaryEntry entry in map) if (string.Equals(Convert.ToString(entry.Key), text, CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase)) {
            object? value = OfficeArgumentValue.Unwrap(entry.Value);
            apply(value is OfficeColor color ? color : OfficeColor.Parse(Convert.ToString(value)!));
            break;
        }
    }

}
