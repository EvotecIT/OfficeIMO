using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using System.Data;

namespace PSWriteOffice.Cmdlets.Excel;

public sealed partial class ExportOfficeExcelCommand {
    private static bool SheetExists(ExcelDocument document, string worksheetName)
    {
        var normalized = NormalizeWorksheetName(worksheetName);
        return document.Sheets.Any(sheet =>
            string.Equals(sheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(sheet.Name, normalized, StringComparison.OrdinalIgnoreCase));
    }

    private static string[] BuildDataSetWorksheetNames(ExcelDocument document, DataSet dataSet, bool allowExistingMatches)
    {
        return BuildDataSetWorksheetNames(dataSet, document.Sheets.Select(sheet => sheet.Name),
            allowExistingMatches ? document : null);
    }

    private static string[] BuildDataSetWorksheetNames(DataSet dataSet, IEnumerable<string> existingNames,
        ExcelDocument? document = null)
    {
        var names = new List<string>(dataSet.Tables.Count);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var existing = new HashSet<string>(existingNames.Where(name => !string.IsNullOrWhiteSpace(name)),
            StringComparer.OrdinalIgnoreCase);
        var tableIndex = 1;
        foreach (DataTable table in dataSet.Tables)
        {
            var requested = string.IsNullOrWhiteSpace(table.TableName)
                ? $"Table{tableIndex}"
                : table.TableName;

            var normalized = NormalizeWorksheetNameInfo(requested);
            var candidate = document != null
                ? FindExistingWorksheetName(document, requested, normalized.Name, used, allowDirectNormalizedMatch: !normalized.UsedDefaultFallback)
                : null;

            string candidateName;
            if (!string.IsNullOrWhiteSpace(candidate) && !used.Contains(candidate!))
            {
                candidateName = candidate!;
            }
            else
            {
                candidateName = normalized.Name;
                var suffix = 2;
                while (used.Contains(candidateName) || existing.Contains(candidateName))
                {
                    candidateName = AppendWorksheetSuffix(normalized.Name, suffix);
                    suffix++;
                }
            }

            used.Add(candidateName);
            names.Add(candidateName);
            tableIndex++;
        }

        return names.ToArray();
    }

    private static string NormalizeWorksheetName(string worksheetName)
    {
        return NormalizeWorksheetNameInfo(worksheetName).Name;
    }

    private static (string Name, bool UsedDefaultFallback) NormalizeWorksheetNameInfo(string worksheetName)
    {
        var baseName = (worksheetName ?? string.Empty).Trim().Trim('\'', ' ');
        var chars = new char[baseName.Length];
        for (var i = 0; i < baseName.Length; i++)
        {
            var ch = baseName[i];
            chars[i] = IsInvalidWorksheetNameCharacter(ch) ? '_' : ch;
        }

        var cleaned = new string(chars).Trim();
        var usedDefaultFallback = false;
        if (string.IsNullOrEmpty(cleaned))
        {
            cleaned = "Sheet1";
            usedDefaultFallback = true;
        }

        var name = cleaned.Length > 31 ? cleaned.Substring(0, 31) : cleaned;
        return (name, usedDefaultFallback);
    }

    private static ExcelSheet GetOrAddDataSetSheet(ExcelDocument document, string sheetName)
    {
        var sheet = document.Sheets.FirstOrDefault(existing =>
            string.Equals(existing.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        return sheet ?? document.AddSheet(sheetName, ExcelSheetNameValidationMode.None);
    }

    private static bool IsInvalidWorksheetNameCharacter(char ch)
    {
        return ch is ':' or '\\' or '/' or '?' or '*' or '[' or ']' ||
            char.IsControl(ch) ||
            char.IsSurrogate(ch);
    }

    private static string? FindExistingWorksheetName(
        ExcelDocument document,
        string requestedName,
        string normalizedName,
        ISet<string> usedNames,
        bool allowDirectNormalizedMatch)
    {
        var existing = document.Sheets
            .Select(sheet => sheet.Name)
            .Where(name => !string.IsNullOrWhiteSpace(name) && !usedNames.Contains(name))
            .ToArray();

        var directMatch = existing.FirstOrDefault(name =>
            string.Equals(name, requestedName, StringComparison.OrdinalIgnoreCase) ||
            (allowDirectNormalizedMatch && string.Equals(name, normalizedName, StringComparison.OrdinalIgnoreCase)));
        if (!string.IsNullOrWhiteSpace(directMatch))
        {
            return directMatch;
        }

        // A lone table named "Budget" does not identify an existing "Budget (2)"
        // sheet. Reuse a suffix only when an earlier table in this DataSet has
        // already claimed the normalized base name.
        if (!usedNames.Contains(normalizedName))
        {
            return null;
        }

        return existing
            .Select(name => (Name: name, Suffix: GetWorksheetNameSuffix(name, normalizedName)))
            .Where(match => match.Suffix.HasValue)
            .OrderBy(match => match.Suffix!.Value)
            .Select(match => match.Name)
            .FirstOrDefault();
    }

    private static string AppendWorksheetSuffix(string worksheetName, int suffix)
    {
        var suffixText = $" ({suffix})";
        var maxBase = 31 - suffixText.Length;
        var basePart = worksheetName.Length > maxBase ? worksheetName.Substring(0, maxBase) : worksheetName;
        return basePart + suffixText;
    }

    private static bool IsSuffixedWorksheetName(string worksheetName, string normalizedName)
    {
        return GetWorksheetNameSuffix(worksheetName, normalizedName).HasValue;
    }

    private static int? GetWorksheetNameSuffix(string worksheetName, string normalizedName)
    {
        var close = worksheetName.LastIndexOf(')');
        var open = worksheetName.LastIndexOf(" (", StringComparison.Ordinal);
        if (close != worksheetName.Length - 1 || open < 0 || open >= close)
        {
            return null;
        }

        var suffixText = worksheetName.Substring(open + 2, close - open - 2);
        if (!int.TryParse(suffixText, out var suffix) || suffix < 2)
        {
            return null;
        }

        return string.Equals(worksheetName, AppendWorksheetSuffix(normalizedName, suffix), StringComparison.OrdinalIgnoreCase)
            ? suffix
            : null;
    }
}
