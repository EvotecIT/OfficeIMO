using OfficeIMO.Excel.Xlsb.Model;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares classic XLSB workbook protection.</summary>
    internal static class XlsbWorkbookProtectionProjector {
        internal const string ElementName = "workbookProtection";

        internal static void Apply(OfficeXmlNode workbook, XlsbWorkbookProtection source) {
            if (source.IsEmpty) return;

            OfficeXmlNode protection = XlsbWorkbookMarkup.Add(workbook, ElementName);
            protection.SetBooleanAttribute("lockStructure", source.LockStructure);
            protection.SetBooleanAttribute("lockWindows", source.LockWindows);
            protection.SetBooleanAttribute("lockRevision", source.LockRevision);
            if (source.WorkbookPassword != 0) {
                protection.SetAttribute("workbookPassword", source.WorkbookPassword.ToString("X4", CultureInfo.InvariantCulture));
            }
            if (source.RevisionsPassword != 0) {
                protection.SetAttribute("revisionsPassword", source.RevisionsPassword.ToString("X4", CultureInfo.InvariantCulture));
            }
        }

        internal static bool Matches(OfficeXmlNode actual, XlsbWorkbookProtection? expected) {
            if (expected?.IsEmpty == true && actual.IsNone) return true;
            if (actual.IsNone || expected == null) return actual.IsNone && expected == null;
            if (actual.FirstChild.Exists || UnsupportedAttribute(actual) != null) return false;

            return TryParsePassword(actual.Attribute("workbookPassword"), out ushort workbookPassword)
                && TryParsePassword(actual.Attribute("revisionsPassword"), out ushort revisionsPassword)
                && workbookPassword == expected.WorkbookPassword
                && revisionsPassword == expected.RevisionsPassword
                && ReadLock(actual, "lockStructure") == expected.LockStructure
                && ReadLock(actual, "lockWindows") == expected.LockWindows
                && ReadLock(actual, "lockRevision") == expected.LockRevision;
        }

        internal static bool ReadLock(OfficeXmlNode protection, string name) =>
            XlsbWorkbookMarkup.ReadBoolean(protection, name) == true;

        internal static bool TryParsePassword(string? value, out ushort password) {
            password = 0;
            return string.IsNullOrWhiteSpace(value)
                || (value!.Length <= 4 && ushort.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out password));
        }

        internal static string? UnsupportedAttribute(OfficeXmlNode protection) {
            foreach (OfficeXmlAttribute attribute in protection.Attributes()) {
                if (attribute.NamespaceUri.Length != 0 || attribute.LocalName is not
                    ("workbookPassword" or "revisionsPassword" or "lockStructure" or "lockWindows" or "lockRevision")) {
                    return attribute.LocalName;
                }
            }
            return null;
        }
    }
}
