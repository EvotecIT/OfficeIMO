using OfficeIMO.GoogleWorkspace;
using System.Text;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsBatchCompiler {
        private static string BuildSheetScopedNamedRangeName(
            ExcelNamedRangeSnapshot namedRange,
            ISet<string> reservedNames,
            TranslationReport report) {
            string sheetPart = SanitizeNamedRangePart(namedRange.SheetName, "Sheet");
            string namePart = SanitizeNamedRangePart(namedRange.Name, "Range");
            string root = sheetPart + "_" + namePart;
            if (!IsIdentifierStart(root[0])) {
                root = "_" + root;
            }

            string target = root;
            for (int suffix = 2; !reservedNames.Add(target); suffix++) {
                target = root + "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            report.Add(
                TranslationSeverity.Info,
                "NamedRanges",
                $"Sheet-scoped Excel name '{namedRange.SheetName}!{namedRange.Name}' was emitted as spreadsheet-scoped Google Sheets name '{target}'.",
                code: "SHEETS.NAMED_RANGE.QUALIFIED",
                action: TranslationAction.Preserve,
                targetId: target);
            return target;
        }

        private static string SanitizeNamedRangePart(string? value, string fallback) {
            var result = new StringBuilder();
            bool lastWasUnderscore = false;
            foreach (char character in value ?? string.Empty) {
                bool allowed = IsAsciiLetter(character) || (character >= '0' && character <= '9');
                char next = allowed ? character : '_';
                if (next == '_' && lastWasUnderscore) {
                    continue;
                }
                result.Append(next);
                lastWasUnderscore = next == '_';
            }
            string sanitized = result.ToString().Trim('_');
            return sanitized.Length == 0 ? fallback : sanitized;
        }

        private static bool IsIdentifierStart(char value) => IsAsciiLetter(value) || value == '_';
        private static bool IsAsciiLetter(char value) => (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');
    }
}
