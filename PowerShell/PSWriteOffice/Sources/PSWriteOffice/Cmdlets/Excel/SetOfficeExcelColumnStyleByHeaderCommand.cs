using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Applies common number, fill, font, and status styles to a worksheet column resolved by header text.</summary>
/// <para>Uses the OfficeIMO header resolver so scripts can style report columns without calculating column letters or ranges.</para>
/// <example>
///   <summary>Format common report columns by header.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' {
///     Set-OfficeExcelColumnStyleByHeader -Header Revenue -Style Currency -CultureName en-US -AutoFit
///     Set-OfficeExcelColumnStyleByHeader -Header Status -BackgroundByText @{ Ready = '#D4EDDA'; Blocked = '#F8D7DA' } -BoldByText Blocked
///   }</code>
///   <para>Styles Revenue as currency and colors Status cells by their text.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelColumnStyleByHeader", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelColumnStyleByHeader", "ExcelColumnStyle")]
public sealed class SetOfficeExcelColumnStyleByHeaderCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>Header caption used to resolve the target column.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Header { get; set; } = string.Empty;

    /// <summary>Include the header cell in the applied formatting.</summary>
    [Parameter]
    public SwitchParameter IncludeHeader { get; set; }

    /// <summary>Preset number style to apply.</summary>
    [Parameter]
    public OfficeExcelColumnStylePreset? Style { get; set; }

    /// <summary>Decimal places for number, percent, and currency styles.</summary>
    [Parameter]
    public int Decimals { get; set; } = 2;

    /// <summary>Culture used by currency formatting, such as en-US or pl-PL.</summary>
    [Parameter]
    public string? CultureName { get; set; }

    /// <summary>Custom number format. Also used when <see cref="Style"/> is NumberFormat.</summary>
    [Parameter]
    public string? NumberFormat { get; set; }

    /// <summary>Date or DateTime number format pattern.</summary>
    [Parameter]
    public string? Pattern { get; set; }

    /// <summary>Apply bold text to the whole resolved column.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }

    /// <summary>Apply a solid background color to the whole resolved column.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }

    /// <summary>Apply a font color to the whole resolved column.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FontColor { get; set; }

    /// <summary>Align cell content in the resolved column.</summary>
    [Parameter]
    public OfficeExcelColumnAlignment? Alignment { get; set; }

    /// <summary>Background colors keyed by matching cell text.</summary>
    [Parameter]
    public Hashtable? BackgroundByText { get; set; }

    /// <summary>Font colors keyed by matching cell text.</summary>
    [Parameter]
    public Hashtable? FontColorByText { get; set; }

    /// <summary>Values that should be bolded when the cell text matches.</summary>
    [Parameter]
    public string[]? BoldByText { get; set; }

    /// <summary>Use case-sensitive matching for text maps.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Set the resolved column width.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Auto-fit the resolved column after applying styles.</summary>
    [Parameter]
    public SwitchParameter AutoFit { get; set; }

    /// <summary>Do nothing when the header cannot be found instead of throwing.</summary>
    [Parameter]
    public SwitchParameter IgnoreMissing { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (string.IsNullOrWhiteSpace(Header)) {
            throw new PSArgumentException("Header cannot be empty.", nameof(Header));
        }

        if (Decimals < 0) {
            throw new ArgumentOutOfRangeException(nameof(Decimals), "Decimals must be zero or greater.");
        }

        var sheet = ResolveSheet();
        bool requiresCellMaterialization = Bold.IsPresent ||
            !string.IsNullOrWhiteSpace(BackgroundColor) ||
            !string.IsNullOrWhiteSpace(FontColor) ||
            Alignment.HasValue ||
            BackgroundByText is { Count: > 0 } ||
            FontColorByText is { Count: > 0 } ||
            BoldByText is { Length: > 0 };

        if (!sheet.TryGetColumnStyleByHeader(
            Header,
            IncludeHeader.IsPresent,
            out var builder,
            out var columnIndex,
            preferDirectTabularMetadata: !requiresCellMaterialization)) {
            if (IgnoreMissing.IsPresent) {
                return;
            }

            throw new PSArgumentException($"Header '{Header}' was not found on worksheet '{sheet.Name}'.", nameof(Header));
        }

        var hasAction = false;

        if (Style.HasValue) {
            ApplyPreset(builder, Style.Value);
            hasAction = true;
        } else if (!string.IsNullOrWhiteSpace(NumberFormat)) {
            builder.NumberFormat(NumberFormat!);
            hasAction = true;
        }

        if (Bold.IsPresent) {
            builder.Bold();
            hasAction = true;
        }

        if (!string.IsNullOrWhiteSpace(BackgroundColor)) {
            builder.Background(BackgroundColor!);
            hasAction = true;
        }

        if (!string.IsNullOrWhiteSpace(FontColor)) {
            builder.FontColor(FontColor!);
            hasAction = true;
        }

        if (Alignment.HasValue) {
            switch (Alignment.Value) {
                case OfficeExcelColumnAlignment.Left:
                    builder.AlignLeft();
                    break;
                case OfficeExcelColumnAlignment.Center:
                    builder.AlignCenter();
                    break;
                case OfficeExcelColumnAlignment.Right:
                    builder.AlignRight();
                    break;
            }
            hasAction = true;
        }

        if (BackgroundByText != null && BackgroundByText.Count > 0) {
            builder.BackgroundByTextMap(ToStringMap(BackgroundByText, CaseSensitive.IsPresent), !CaseSensitive.IsPresent);
            hasAction = true;
        }

        if (FontColorByText != null && FontColorByText.Count > 0) {
            builder.FontColorByTextMap(ToStringMap(FontColorByText, CaseSensitive.IsPresent), !CaseSensitive.IsPresent);
            hasAction = true;
        }

        if (BoldByText != null && BoldByText.Length > 0) {
            var comparer = CaseSensitive.IsPresent ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            builder.BoldByTextSet(new HashSet<string>(BoldByText, comparer), !CaseSensitive.IsPresent);
            hasAction = true;
        }

        if (Width.HasValue) {
            sheet.SetColumnWidth(columnIndex, Width.Value);
            hasAction = true;
        }

        if (AutoFit.IsPresent) {
            sheet.AutoFitColumn(columnIndex);
            hasAction = true;
        }

        if (!hasAction) {
            throw new PSArgumentException("Provide a style, color, width, AutoFit, or text map option to update the column.");
        }
        WritePassThru(sheet);
    }

    private ExcelSheet ResolveSheet() {
        if (ParameterSetName == ParameterSetDocument) {
            if (Document == null) {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        return ExcelDslContext.Require(this).RequireSheet();
    }

    private void ApplyPreset(ExcelColumnStyleByHeaderBuilder builder, OfficeExcelColumnStylePreset style) {
        switch (style) {
            case OfficeExcelColumnStylePreset.Number:
                builder.Number(Decimals);
                break;
            case OfficeExcelColumnStylePreset.Integer:
                builder.Integer();
                break;
            case OfficeExcelColumnStylePreset.Percent:
                builder.Percent(Decimals);
                break;
            case OfficeExcelColumnStylePreset.Currency:
                builder.Currency(Decimals, ResolveCulture());
                break;
            case OfficeExcelColumnStylePreset.Date:
                builder.Date(string.IsNullOrWhiteSpace(Pattern) ? "yyyy-mm-dd" : Pattern!);
                break;
            case OfficeExcelColumnStylePreset.DateTime:
                builder.DateTime(string.IsNullOrWhiteSpace(Pattern) ? "yyyy-mm-dd hh:mm:ss" : Pattern!);
                break;
            case OfficeExcelColumnStylePreset.Time:
                builder.Time();
                break;
            case OfficeExcelColumnStylePreset.DurationHours:
                builder.DurationHours();
                break;
            case OfficeExcelColumnStylePreset.Text:
                builder.Text();
                break;
            case OfficeExcelColumnStylePreset.NumberFormat:
                if (string.IsNullOrWhiteSpace(NumberFormat)) {
                    throw new PSArgumentException("Provide -NumberFormat when -Style NumberFormat is used.", nameof(NumberFormat));
                }
                builder.NumberFormat(NumberFormat!);
                break;
        }
    }

    private CultureInfo? ResolveCulture() {
        if (string.IsNullOrWhiteSpace(CultureName)) {
            return null;
        }

        return CultureInfo.GetCultureInfo(CultureName!);
    }

    private static Dictionary<string, string> ToStringMap(Hashtable table, bool caseSensitive) {
        var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        var result = new Dictionary<string, string>(comparer);
        foreach (DictionaryEntry entry in table) {
            var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            var value = Convert.ToString(entry.Value, CultureInfo.InvariantCulture);
            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value)) {
                result[key] = value;
            }
        }

        return result;
    }
}
