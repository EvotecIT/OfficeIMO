using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a title block to the current Excel report sheet.</summary>
/// <example>
///   <summary>Add a polished title to a report sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\Operations.xlsx {
///     Add-OfficeExcelReportSheet -Name Summary {
///         Add-OfficeExcelReportTitle -Title 'Operational Summary' -Subtitle 'Current month'
///         Add-OfficeExcelReportKpiRow -InputObject @{ Revenue = 125000; Incidents = 3; Status = 'Ready' }
///     }
/// }</code>
///   <para>Uses the OfficeIMO sheet composer through PSWriteOffice's thin report-block wrapper.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportTitle")]
[Alias("ExcelReportTitle")]
public sealed class AddOfficeExcelReportTitleCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Title text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Optional subtitle text.</summary>
    [Parameter(Position = 1)]
    public string? Subtitle { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().Title(Title, Subtitle);
        WritePassThru(context.RequireSheet());
    }
}

/// <summary>Adds a section heading to the current Excel report sheet.</summary>
/// <example>
///   <summary>Add a section and paragraph to a report sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\Operations.xlsx {
///     Add-OfficeExcelReportSheet -Name Summary {
///         Add-OfficeExcelReportSection -Text 'Service health'
///         Add-OfficeExcelReportParagraph -Text 'All monitored services are reporting.'
///     }
/// }</code>
///   <para>Uses the report composer to add a section heading and narrative text.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportSection")]
[Alias("ExcelReportSection")]
public sealed class AddOfficeExcelReportSectionCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Section heading text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().Section(Text);
        WritePassThru(context.RequireSheet());
    }
}

/// <summary>Adds a paragraph line to the current Excel report sheet.</summary>
/// <example>
///   <summary>Add narrative text below a title block.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\Operations.xlsx {
///     Add-OfficeExcelReportSheet -Name Summary {
///         Add-OfficeExcelReportTitle -Title 'Operational Summary'
///         Add-OfficeExcelReportParagraph -Text 'This workbook was generated from the validated source data.'
///     }
/// }</code>
///   <para>Adds prose to an OfficeIMO-composed Excel report sheet.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportParagraph")]
[Alias("ExcelReportParagraph")]
public sealed class AddOfficeExcelReportParagraphCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Paragraph text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().Paragraph(Text);
        WritePassThru(context.RequireSheet());
    }
}

/// <summary>Adds vertical spacing to the current Excel report sheet.</summary>
/// <example>
///   <summary>Add spacing between report blocks.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\Operations.xlsx {
///     Add-OfficeExcelReportSheet -Name Summary {
///         Add-OfficeExcelReportTitle -Title 'Operational Summary'
///         Add-OfficeExcelReportSpacer -Rows 2
///         Add-OfficeExcelReportSection -Text 'Details'
///     }
/// }</code>
///   <para>Advances the composer cursor before adding the next report block.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportSpacer")]
[Alias("ExcelReportSpacer")]
public sealed class AddOfficeExcelReportSpacerCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Rows to advance. Defaults to the composer theme spacing.</summary>
    [Parameter(Position = 0)]
    public int Rows { get; set; } = -1;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().Spacer(Rows);
        WritePassThru(context.RequireSheet());
    }
}

/// <summary>Adds a colored callout block to the current Excel report sheet.</summary>
/// <example>
///   <summary>Add a warning callout to a report sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\Operations.xlsx {
///     Add-OfficeExcelReportSheet -Name Summary {
///         Add-OfficeExcelReportCallout -Kind Warning -Title 'Manual validation' -Body 'Open the workbook in desktop Excel before publishing pivot-heavy reports.'
///     }
/// }</code>
///   <para>Renders a composer callout block using the current report theme.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportCallout")]
[Alias("ExcelReportCallout")]
public sealed class AddOfficeExcelReportCalloutCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Callout kind. Supported values include info, success, warning, error, and critical.</summary>
    [Parameter(Position = 0)]
    public OfficeExcelReportCalloutKind Kind { get; set; } = OfficeExcelReportCalloutKind.Info;

    /// <summary>Callout title.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Callout body text.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public string Body { get; set; } = string.Empty;

    /// <summary>Width of the highlighted callout band in worksheet columns.</summary>
    [Parameter]
    public int WidthColumns { get; set; } = 8;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().Callout(Kind.ToString(), Title, Body, WidthColumns);
        WritePassThru(context.RequireSheet());
    }
}

/// <summary>Adds a KPI row to the current Excel report sheet.</summary>
/// <example>
///   <summary>Add three KPI values to a summary sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\Operations.xlsx {
///     Add-OfficeExcelReportSheet -Name Summary {
///         Add-OfficeExcelReportKpiRow -InputObject @{ Revenue = 125000; Incidents = 3; Status = 'Ready' } -PerRow 3
///     }
/// }</code>
///   <para>Renders PowerShell key/value data as a KPI row through the OfficeIMO sheet composer.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportKpiRow")]
[Alias("ExcelReportKpiRow")]
public sealed class AddOfficeExcelReportKpiRowCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Hashtable or objects with Label/Value, Key/Value, Name/Value, or Title/Value properties.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("Data")]
    public object InputObject { get; set; } = null!;

    /// <summary>Number of KPI cards per rendered row.</summary>
    [Parameter]
    public int PerRow { get; set; } = 3;

    /// <summary>Optional fill color for KPI labels.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LabelFillColor { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().KpiRow(ReportBlockInput.ToPairs(InputObject), PerRow, LabelFillColor);
        WritePassThru(context.RequireSheet());
    }
}

/// <summary>Adds a legend table to the current Excel report sheet.</summary>
/// <example>
///   <summary>Add a status legend with colored first-column values.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$legendRows = @(
///     [pscustomobject]@{ Status = 'Ready'; Meaning = 'Validated and ready' }
///     [pscustomobject]@{ Status = 'Review'; Meaning = 'Needs owner review' }
/// )
/// New-OfficeExcel -Path .\Operations.xlsx {
///     Add-OfficeExcelReportSheet -Name Summary {
///         Add-OfficeExcelReportLegend -Title 'Status legend' -Header Status, Meaning -InputObject $legendRows -FirstColumnFillByValue @{ Ready = '#d9f7be'; Review = '#fff7e6' }
///     }
/// }</code>
///   <para>Renders legend rows and applies optional fill colors keyed by the first column.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportLegend")]
[Alias("ExcelReportLegend")]
public sealed class AddOfficeExcelReportLegendCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Optional legend title.</summary>
    [Parameter(Position = 0)]
    public string? Title { get; set; }

    /// <summary>Column headers.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Headers")]
    public string[] Header { get; set; } = Array.Empty<string>();

    /// <summary>Rows. Each row may be an array, enumerable, hashtable, or object.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Rows")]
    public object[] InputObject { get; set; } = Array.Empty<object>();

    /// <summary>Optional first-column fill colors keyed by first-column value.</summary>
    [Parameter]
    public Hashtable? FirstColumnFillByValue { get; set; }

    /// <summary>Optional header fill color.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? HeaderFillColor { get; set; }

    /// <summary>Use case-sensitive matching for first-column fill values.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().SectionLegend(
            Title,
            Header,
            InputObject.Select(row => ReportBlockInput.ToRow(row, Header)),
            ReportBlockInput.ToStringMap(FirstColumnFillByValue, CaseSensitive.IsPresent),
            HeaderFillColor);
        WritePassThru(context.RequireSheet());
    }
}

/// <summary>Adds an object table to the current Excel report sheet using the OfficeIMO sheet composer.</summary>
/// <example>
///   <summary>Add a styled report table from objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @(
///     [pscustomobject]@{ Area = 'PDF'; Status = 'Ready' }
///     [pscustomobject]@{ Area = 'Word'; Status = 'Review' }
/// )
/// New-OfficeExcel -Path .\Operations.xlsx {
///     Add-OfficeExcelReportSheet -Name Summary {
///         Add-OfficeExcelReportTable -InputObject $rows -Title 'Documentation coverage' -TableStyle TableStyleMedium9
///     }
/// }</code>
///   <para>Renders object rows as a formatted Excel table through the sheet composer.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportTable")]
[Alias("ExcelReportTable")]
public sealed class AddOfficeExcelReportTableCommand : PSCmdlet {
    private readonly List<object?> _items = new();

    /// <summary>Objects to flatten and render as a table.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("Data")]
    public object? InputObject { get; set; }

    /// <summary>Optional section title displayed above the table.</summary>
    [Parameter(Position = 1)]
    public string? Title { get; set; }

    /// <summary>Built-in table style to apply.</summary>
    [Parameter]
    public ExcelTableStyle TableStyle { get; set; } = ExcelTableStyle.TableStyleMedium9;

    /// <summary>Properties to include, in the requested column order.</summary>
    [Parameter]
    public string[]? Property { get; set; }

    /// <summary>Properties to exclude from the rendered table.</summary>
    [Parameter]
    public string[]? ExcludeProperty { get; set; }

    /// <summary>Text used between items when a cell contains a collection.</summary>
    [Parameter]
    [AllowEmptyString]
    public string CollectionSeparator { get; set; } = ", ";

    /// <summary>Text used between entries when a cell contains a dictionary.</summary>
    [Parameter]
    [AllowEmptyString]
    public string DictionaryEntrySeparator { get; set; } = "; ";

    /// <summary>Text used between a dictionary key and value.</summary>
    [Parameter]
    [AllowEmptyString]
    public string DictionaryKeyValueSeparator { get; set; } = ": ";

    /// <summary>Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxCollectionItems { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems;

    /// <summary>Maximum nesting depth allowed while normalizing one cell value. Defaults to 64; increase explicitly for trusted deeper values.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxNestingDepth { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth;

    /// <summary>Emphasize the first table column when the selected style supports it.</summary>
    [Parameter]
    public SwitchParameter ShowFirstColumn { get; set; }

    /// <summary>Emphasize the last table column when the selected style supports it.</summary>
    [Parameter]
    public SwitchParameter ShowLastColumn { get; set; }

    /// <summary>Disable alternating row stripes for the generated table.</summary>
    [Parameter]
    public SwitchParameter NoRowStripes { get; set; }

    /// <summary>Enable alternating column stripes for the generated table.</summary>
    [Parameter]
    public SwitchParameter ShowColumnStripes { get; set; }

    /// <summary>Disable AutoFilter dropdowns.</summary>
    [Parameter]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Do not freeze through the table header row.</summary>
    [Parameter]
    public SwitchParameter NoFreezeHeaderRow { get; set; }

    /// <summary>Disable composer auto-formatting for dynamic collection columns.</summary>
    [Parameter]
    public SwitchParameter NoAutoFormatDynamicCollections { get; set; }

    /// <summary>Emit the A1 range used by the generated table.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TableInputCollector.AddInput(_items, InputObject, preserveTabularInput: true);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        var inputRows = TableInputCollector.RequireRows(_items, nameof(InputObject));
        var normalizerOptions = PowerShellObjectNormalizerOptions.ForTable(
            CollectionSeparator,
            DictionaryEntrySeparator,
            DictionaryKeyValueSeparator,
            MaxCollectionItems,
            MaxNestingDepth);
        var table = ExcelTabularInputService.ToDataTable(
            inputRows,
            copyExistingTables: false,
            normalizerOptions: normalizerOptions);
        var sourceColumns = table.Columns
            .Cast<System.Data.DataColumn>()
            .Select(static column => column.ColumnName)
            .ToArray();
        var requestedColumns = NormalizePropertyList(Property) ?? sourceColumns;
        var excludedProperties = NormalizePropertyList(ExcludeProperty) ?? Array.Empty<string>();

        var style = TableStyle;

        var composer = ExcelDslContext.Require(this).RequireComposer();
        var range = composer.TableFrom(
            table,
            Title,
            configure: options => {
                options.Columns = requestedColumns;
                options.ExcludeProperties = excludedProperties;
            },
            style: style,
            autoFilter: !NoAutoFilter.IsPresent,
            freezeHeaderRow: !NoFreezeHeaderRow.IsPresent,
            visuals: options => {
                options.AutoFormatDynamicCollections = !NoAutoFormatDynamicCollections.IsPresent;
                options.ShowFirstColumn = ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowFirstColumn), ShowFirstColumn);
                options.ShowLastColumn = ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowLastColumn), ShowLastColumn);
                options.ShowRowStripes = !ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(NoRowStripes), NoRowStripes);
                options.ShowColumnStripes = ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowColumnStripes), ShowColumnStripes);
            });

        if (PassThru.IsPresent) {
            WriteObject(range);
        }
    }

    private static string[]? NormalizePropertyList(string[]? properties) {
        if (properties == null || properties.Length == 0) {
            return null;
        }

        var normalized = properties
            .Where(static property => !string.IsNullOrWhiteSpace(property))
            .Select(static property => property.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        return normalized.Length == 0 ? null : normalized;
    }
}

internal static class ReportBlockInput {
    public static IReadOnlyList<(string Label, object? Value)> ToPairs(object input) {
        if (input is IDictionary dictionary) {
            var pairs = new List<(string Label, object? Value)>();
            foreach (DictionaryEntry entry in dictionary) {
                pairs.Add((Convert.ToString(entry.Key, CultureInfo.InvariantCulture) ?? string.Empty, entry.Value));
            }

            return pairs;
        }

        var rows = input is IEnumerable enumerable and not string
            ? enumerable.Cast<object?>().Where(item => item != null).ToArray()
            : new[] { input };

        return rows.Select(item => {
            var psObject = PSObject.AsPSObject(item);
            var label = GetProperty(psObject, "Label")
                ?? GetProperty(psObject, "Key")
                ?? GetProperty(psObject, "Name")
                ?? GetProperty(psObject, "Title")
                ?? Convert.ToString(item, CultureInfo.InvariantCulture)
                ?? string.Empty;
            var value = GetPropertyValue(psObject, "Value") ?? GetPropertyValue(psObject, "Count") ?? GetPropertyValue(psObject, "Total");
            return (label, value);
        }).ToArray();
    }

    public static IReadOnlyList<string> ToRow(object input, IReadOnlyList<string> headers) {
        if (input is IDictionary dictionary) {
            return headers.Select(header => {
                foreach (DictionaryEntry entry in dictionary) {
                    if (string.Equals(Convert.ToString(entry.Key, CultureInfo.InvariantCulture), header, StringComparison.OrdinalIgnoreCase)) {
                        return Convert.ToString(entry.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                    }
                }

                return string.Empty;
            }).ToArray();
        }

        if (input is IEnumerable enumerable and not string) {
            return enumerable.Cast<object?>()
                .Select(value => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
                .ToArray();
        }

        var psObject = PSObject.AsPSObject(input);
        return headers.Select(header => Convert.ToString(GetPropertyValue(psObject, header), CultureInfo.InvariantCulture) ?? string.Empty).ToArray();
    }

    public static Dictionary<string, string>? ToStringMap(Hashtable? table, bool caseSensitive = false) {
        if (table == null || table.Count == 0) {
            return null;
        }

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

    private static string? GetProperty(PSObject psObject, string name) {
        return Convert.ToString(GetPropertyValue(psObject, name), CultureInfo.InvariantCulture);
    }

    private static object? GetPropertyValue(PSObject psObject, string name) {
        return psObject.Properties[name]?.Value;
    }
}
