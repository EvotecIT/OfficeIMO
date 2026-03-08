using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

internal static class SheetNameLookup
{
    private static readonly Regex MultipleUnderscoresRegex = new("_+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    internal static ExcelSheet? FindByRequestedName(IEnumerable<ExcelSheet> sheets, string requestedName)
    {
        return sheets.FirstOrDefault(sheet => Matches(sheet.Name, requestedName));
    }

    internal static Sheet? FindByRequestedName(IEnumerable<Sheet> sheets, string requestedName)
    {
        return sheets.FirstOrDefault(sheet => Matches(sheet.Name?.Value, requestedName));
    }

    internal static string ResolveExistingOrRequested(IEnumerable<ExcelSheet> sheets, string requestedName)
    {
        return FindByRequestedName(sheets, requestedName)?.Name ?? requestedName;
    }

    internal static string ResolveExistingOrNormalizedOrRequested(IEnumerable<ExcelSheet> sheets, string requestedName)
    {
        return FindByRequestedName(sheets, requestedName)?.Name
               ?? NormalizeForLookup(requestedName)
               ?? requestedName;
    }

    internal static string BuildInternalLocation(IEnumerable<ExcelSheet> sheets, string requestedName, string targetA1, bool normalizeFallback = false)
    {
        string effectiveSheetName = normalizeFallback
            ? ResolveExistingOrNormalizedOrRequested(sheets, requestedName)
            : ResolveExistingOrRequested(sheets, requestedName);
        return $"'{ExcelSheet.EscapeSheetNameForLink(effectiveSheetName)}'!{targetA1}";
    }

    internal static bool Matches(string? actualSheetName, string requestedName)
    {
        if (string.Equals(actualSheetName, requestedName, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        string? normalizedRequestedName = NormalizeForLookup(requestedName);
        return normalizedRequestedName != null
               && !string.Equals(normalizedRequestedName, requestedName, StringComparison.OrdinalIgnoreCase)
               && string.Equals(actualSheetName, normalizedRequestedName, StringComparison.OrdinalIgnoreCase);
    }

    internal static string? NormalizeForLookup(string? sheetName)
    {
        string baseName = (sheetName ?? string.Empty).Trim();
        baseName = baseName.Trim('\'', ' ');
        if (baseName.Length == 0)
        {
            return null;
        }

        var sb = new StringBuilder(baseName.Length);
        foreach (char c in baseName)
        {
            if (c == ':' || c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']')
            {
                sb.Append('_');
            }
            else
            {
                sb.Append(c);
            }
        }

        string cleaned = sb.ToString().Trim();
        cleaned = MultipleUnderscoresRegex.Replace(cleaned, "_");
        cleaned = cleaned.Trim('_');
        if (cleaned.Length == 0)
        {
            return null;
        }

        return cleaned.Length > 31 ? cleaned.Substring(0, 31) : cleaned;
    }
}
