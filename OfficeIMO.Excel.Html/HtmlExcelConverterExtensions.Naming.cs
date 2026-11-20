using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static string GetSheetName(IElement section) {
        string? name = section.GetAttribute("data-officeimo-sheet");
        if (!string.IsNullOrWhiteSpace(name)) {
            return name!.Trim();
        }

        return NormalizeText(section.QuerySelector("h2")?.TextContent) is { Length: > 0 } heading
            ? heading
            : "Sheet";
    }

    private static string GetUniqueSheetName(string name, HashSet<string> usedNames) {
        string baseName = SanitizeSheetName(name);
        string candidate = baseName;
        int suffix = 2;
        while (!usedNames.Add(candidate)) {
            string suffixText = " " + suffix.ToString(CultureInfo.InvariantCulture);
            int maxBaseLength = Math.Max(1, 31 - suffixText.Length);
            candidate = ExcelSheetNameText.Truncate(baseName, maxBaseLength) + suffixText;
            suffix++;
        }

        return candidate;
    }

    private static string SanitizeSheetName(string name) {
        string value = string.IsNullOrWhiteSpace(name) ? "Sheet" : name.Trim();
        foreach (char invalid in new[] { ':', '\\', '/', '?', '*', '[', ']' }) {
            value = value.Replace(invalid, '-');
        }

        return ExcelSheetNameText.Truncate(value, 31);
    }
}
