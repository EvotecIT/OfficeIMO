using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds an AutoFilter to the current worksheet.</summary>
/// <para>Optional filter criteria can be supplied per column index (0-based within the range).</para>
/// <example>
///   <summary>Add AutoFilter to a range.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelAutoFilter -Range 'A1:D200' }</code>
///   <para>Enables filter dropdowns on the range.</para>
/// </example>
/// <example>
///   <summary>Add AutoFilter with criteria.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeExcelAutoFilter -Range 'A1:D200' -Criteria @{ 2 = 'Open','Hold' }</code>
///   <para>Filters the third column (0-based within the range) to Open/Hold.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelAutoFilter", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelAutoFilter")]
public sealed class AddOfficeExcelAutoFilterCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
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

    /// <summary>A1 range to apply AutoFilter.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Optional criteria per column index (0-based within the range).</summary>
    [Parameter]
    public Hashtable? Criteria { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ResolveSheet();
        var criteria = ConvertCriteria(Criteria);
        sheet.AddAutoFilter(Range, criteria);
        WritePassThru(sheet);
    }

    private ExcelSheet ResolveSheet() {
        if (ParameterSetName == ParameterSetDocument) {
            if (Document == null) {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }

    private static Dictionary<uint, IEnumerable<string>>? ConvertCriteria(Hashtable? criteria) {
        if (criteria == null || criteria.Count == 0) {
            return null;
        }

        var converted = new Dictionary<uint, IEnumerable<string>>();
        foreach (DictionaryEntry entry in criteria) {
            if (entry.Key == null) {
                continue;
            }

            if (!TryGetColumnIndex(entry.Key, out uint index)) {
                throw new PSArgumentException($"Invalid column index '{entry.Key}'. Use an integer index (0-based within the filter range).");
            }

            var values = NormalizeValues(entry.Value).ToArray();
            if (values.Length == 0) {
                continue;
            }

            converted[index] = values;
        }

        return converted.Count == 0 ? null : converted;
    }

    private static bool TryGetColumnIndex(object key, out uint index) {
        index = 0;
        switch (key) {
            case byte b:
                index = b;
                return true;
            case sbyte sb when sb >= 0:
                index = (uint)sb;
                return true;
            case short s when s >= 0:
                index = (uint)s;
                return true;
            case ushort us:
                index = us;
                return true;
            case int i when i >= 0:
                index = (uint)i;
                return true;
            case uint ui:
                index = ui;
                return true;
            case long l when l >= 0 && l <= uint.MaxValue:
                index = (uint)l;
                return true;
            case ulong ul when ul <= uint.MaxValue:
                index = (uint)ul;
                return true;
            case string text:
                return uint.TryParse(text, out index);
            default:
                return false;
        }
    }

    private static IEnumerable<string> NormalizeValues(object? value) {
        if (value == null) {
            yield break;
        }

        if (value is string text) {
            if (!string.IsNullOrWhiteSpace(text)) {
                yield return text;
            }
            yield break;
        }

        if (value is IEnumerable enumerable) {
            foreach (var item in enumerable) {
                if (item == null) continue;
                var itemText = item.ToString();
                if (!string.IsNullOrWhiteSpace(itemText)) {
                    yield return itemText!;
                }
            }
            yield break;
        }

        var single = value.ToString();
        if (!string.IsNullOrWhiteSpace(single)) {
            yield return single!;
        }
    }
}