using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares classic XLSB worksheet protection.</summary>
    internal static class XlsbWorksheetProtectionProjector {
        internal static void Apply(ExcelSheet sheet, XlsbWorksheetProtection? source) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (source == null) return;

            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            worksheet.Append(Create(source));
        }

        internal static void ValidateUnchanged(ExcelSheet sheet, XlsbWorksheetProtection? expected) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            SheetProtection[] actual = worksheet.Elements<SheetProtection>().ToArray();
            if (actual.Length > 1
                || (expected == null && actual.Length != 0)
                || (expected != null && (actual.Length != 1 || !Matches(actual[0], expected)))) {
                throw new NotSupportedException($"Native XLSB rewriting preserves but cannot modify worksheet protection on worksheet '{sheet.Name}'. Save as .xlsx to retain that change.");
            }
        }

        internal static bool TryParsePassword(string? value, out ushort password) {
            password = 0;
            return string.IsNullOrWhiteSpace(value)
                || (value!.Length <= 4 && ushort.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out password));
        }

        private static SheetProtection Create(XlsbWorksheetProtection source) {
            var protection = new SheetProtection {
                Sheet = source.IsProtected,
                Objects = !source.AllowEditObjects,
                Scenarios = !source.AllowEditScenarios,
                FormatCells = !source.AllowFormatCells,
                FormatColumns = !source.AllowFormatColumns,
                FormatRows = !source.AllowFormatRows,
                InsertColumns = !source.AllowInsertColumns,
                InsertRows = !source.AllowInsertRows,
                InsertHyperlinks = !source.AllowInsertHyperlinks,
                DeleteColumns = !source.AllowDeleteColumns,
                DeleteRows = !source.AllowDeleteRows,
                SelectLockedCells = !source.AllowSelectLockedCells,
                Sort = !source.AllowSort,
                AutoFilter = !source.AllowAutoFilter,
                PivotTables = !source.AllowPivotTables,
                SelectUnlockedCells = !source.AllowSelectUnlockedCells
            };
            if (source.Password != 0) {
                protection.Password = source.Password.ToString("X4", CultureInfo.InvariantCulture);
            }
            return protection;
        }

        private static bool Matches(SheetProtection actual, XlsbWorksheetProtection expected) {
            if (actual.HasChildren
                || actual.GetAttributes().Any(attribute =>
                    !string.IsNullOrEmpty(attribute.NamespaceUri)
                    || !IsSupportedAttribute(attribute.LocalName))
                || !TryParsePassword(actual.Password?.Value, out ushort password)) {
                return false;
            }

            return password == expected.Password
                && (actual.Sheet?.Value ?? true) == expected.IsProtected
                && IsAllowed(actual.Objects, lockedWhenOmitted: false) == expected.AllowEditObjects
                && IsAllowed(actual.Scenarios, lockedWhenOmitted: false) == expected.AllowEditScenarios
                && IsAllowed(actual.FormatCells, lockedWhenOmitted: true) == expected.AllowFormatCells
                && IsAllowed(actual.FormatColumns, lockedWhenOmitted: true) == expected.AllowFormatColumns
                && IsAllowed(actual.FormatRows, lockedWhenOmitted: true) == expected.AllowFormatRows
                && IsAllowed(actual.InsertColumns, lockedWhenOmitted: true) == expected.AllowInsertColumns
                && IsAllowed(actual.InsertRows, lockedWhenOmitted: true) == expected.AllowInsertRows
                && IsAllowed(actual.InsertHyperlinks, lockedWhenOmitted: true) == expected.AllowInsertHyperlinks
                && IsAllowed(actual.DeleteColumns, lockedWhenOmitted: true) == expected.AllowDeleteColumns
                && IsAllowed(actual.DeleteRows, lockedWhenOmitted: true) == expected.AllowDeleteRows
                && IsAllowed(actual.SelectLockedCells, lockedWhenOmitted: false) == expected.AllowSelectLockedCells
                && IsAllowed(actual.Sort, lockedWhenOmitted: true) == expected.AllowSort
                && IsAllowed(actual.AutoFilter, lockedWhenOmitted: true) == expected.AllowAutoFilter
                && IsAllowed(actual.PivotTables, lockedWhenOmitted: true) == expected.AllowPivotTables
                && IsAllowed(actual.SelectUnlockedCells, lockedWhenOmitted: false) == expected.AllowSelectUnlockedCells;
        }

        private static bool IsAllowed(BooleanValue? protectionFlag, bool lockedWhenOmitted) =>
            !(protectionFlag?.Value ?? lockedWhenOmitted);

        private static bool IsSupportedAttribute(string name) =>
            name == "password"
            || name == "sheet"
            || name == "objects"
            || name == "scenarios"
            || name == "formatCells"
            || name == "formatColumns"
            || name == "formatRows"
            || name == "insertColumns"
            || name == "insertRows"
            || name == "insertHyperlinks"
            || name == "deleteColumns"
            || name == "deleteRows"
            || name == "selectLockedCells"
            || name == "sort"
            || name == "autoFilter"
            || name == "pivotTables"
            || name == "selectUnlockedCells";
    }
}
