using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares classic XLSB worksheet protection.</summary>
    internal static class XlsbWorksheetProtectionProjector {
        internal const string ElementName = "sheetProtection";

        // The order of the permissions is also the order of the fields in BrtSheetProtection.
        private static readonly (string Name, bool LockedByDefault, Func<XlsbWorksheetProtection, bool> Allow)[] Permissions = {
            ("objects", false, source => source.AllowEditObjects),
            ("scenarios", false, source => source.AllowEditScenarios),
            ("formatCells", true, source => source.AllowFormatCells),
            ("formatColumns", true, source => source.AllowFormatColumns),
            ("formatRows", true, source => source.AllowFormatRows),
            ("insertColumns", true, source => source.AllowInsertColumns),
            ("insertRows", true, source => source.AllowInsertRows),
            ("insertHyperlinks", true, source => source.AllowInsertHyperlinks),
            ("deleteColumns", true, source => source.AllowDeleteColumns),
            ("deleteRows", true, source => source.AllowDeleteRows),
            ("selectLockedCells", false, source => source.AllowSelectLockedCells),
            ("sort", true, source => source.AllowSort),
            ("autoFilter", true, source => source.AllowAutoFilter),
            ("pivotTables", true, source => source.AllowPivotTables),
            ("selectUnlockedCells", false, source => source.AllowSelectUnlockedCells)
        };

        private static readonly string[] SupportedAttributes = Permissions.Select(permission => permission.Name)
            .Concat(new[] { "password", "sheet" }).ToArray();

        internal static void Apply(ExcelSheet sheet, XlsbWorksheetProtection? source) {
            if (source != null) Apply(XlsbWorksheetMarkup.Root(sheet), source);
        }

        internal static void Apply(OfficeXmlNode worksheet, XlsbWorksheetProtection source) {
            OfficeXmlNode protection = XlsbWorksheetMarkup.Append(worksheet, ElementName);
            protection.SetBooleanAttribute("sheet", source.IsProtected);
            foreach (var permission in Permissions) protection.SetBooleanAttribute(permission.Name, !permission.Allow(source));
            if (source.Password != 0) protection.SetAttribute("password", source.Password.ToString("X4", CultureInfo.InvariantCulture));
        }

        internal static void ValidateUnchanged(ExcelSheet sheet, XlsbWorksheetProtection? expected) {
            OfficeXmlNode worksheet = XlsbWorksheetMarkup.Root(sheet);
            var children = worksheet.Children(OfficeXmlNamespaces.SpreadsheetML, ElementName).GetEnumerator();
            if (!children.MoveNext()) {
                if (expected == null) return;
            } else if (expected != null) {
                OfficeXmlNode actual = children.Current;
                if (!children.MoveNext() && Matches(actual, expected)) return;
            }
            throw new NotSupportedException($"Native XLSB rewriting preserves but cannot modify worksheet protection on worksheet '{sheet.Name}'. Save as .xlsx to retain that change.");
        }

        private static bool Matches(OfficeXmlNode actual, XlsbWorksheetProtection expected) {
            if (actual.FirstChild.Exists || UnsupportedAttribute(actual) != null
                || !XlsbWorkbookProtectionProjector.TryParsePassword(actual.Attribute("password"), out ushort password)
                || password != expected.Password || IsProtected(actual) != expected.IsProtected) return false;
            foreach (var permission in Permissions) {
                if (IsAllowed(actual, permission.Name, permission.LockedByDefault) != permission.Allow(expected)) return false;
            }
            return true;
        }

        internal static string? UnsupportedAttribute(OfficeXmlNode protection) =>
            XlsbWorkbookMarkup.UnsupportedAttribute(protection, SupportedAttributes);

        internal static bool IsProtected(OfficeXmlNode protection) =>
            XlsbWorkbookMarkup.ReadBoolean(protection, "sheet") ?? true;

        /// <summary>Yields the permission bits the native worksheet protection record stores.</summary>
        internal static IEnumerable<bool> AllowedActions(OfficeXmlNode protection) {
            foreach (var permission in Permissions) yield return IsAllowed(protection, permission.Name, permission.LockedByDefault);
        }

        private static bool IsAllowed(OfficeXmlNode protection, string name, bool lockedByDefault) =>
            !(XlsbWorkbookMarkup.ReadBoolean(protection, name) ?? lockedByDefault);
    }
}
