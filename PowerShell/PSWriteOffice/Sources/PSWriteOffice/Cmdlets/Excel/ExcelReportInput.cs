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
using PSWriteOffice.Services.Text;
namespace PSWriteOffice.Cmdlets.Excel;

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

    public static Dictionary<string, OfficeIMO.OfficeColor>? ToColorMap(Hashtable? table, bool caseSensitive = false) {
        if (table == null || table.Count == 0) {
            return null;
        }

        var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        var result = new Dictionary<string, OfficeIMO.OfficeColor>(comparer);
        foreach (DictionaryEntry entry in table) {
            var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            var value = OfficeArgumentValue.Unwrap(entry.Value);
            if (!string.IsNullOrEmpty(key) && value != null) {
                OfficeColor? color = value is OfficeColor native ? native : OfficeColor.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!);
                if (color.HasValue) result[key] = color.Value;
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
