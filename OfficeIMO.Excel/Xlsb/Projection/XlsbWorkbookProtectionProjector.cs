using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares classic XLSB workbook protection.</summary>
    internal static class XlsbWorkbookProtectionProjector {
        internal static void Apply(ExcelDocument document, XlsbWorkbookProtection source) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (source.IsEmpty) return;

            var protection = new WorkbookProtection {
                LockStructure = source.LockStructure,
                LockWindows = source.LockWindows,
                LockRevision = source.LockRevision
            };
            if (source.WorkbookPassword != 0) {
                protection.WorkbookPassword = source.WorkbookPassword.ToString("X4", CultureInfo.InvariantCulture);
            }
            if (source.RevisionsPassword != 0) {
                protection.RevisionsPassword = source.RevisionsPassword.ToString("X4", CultureInfo.InvariantCulture);
            }
            OpenXmlWorkbookElementOrder.InsertInOrder(document.WorkbookRoot, protection);
        }

        internal static bool Matches(WorkbookProtection? actual, XlsbWorkbookProtection? expected) {
            if (expected?.IsEmpty == true && actual == null) return true;
            if (actual == null || expected == null || actual.HasChildren) return actual == null && expected == null;
            if (actual.GetAttributes().Any(attribute =>
                    !string.IsNullOrEmpty(attribute.NamespaceUri)
                    || !IsSupportedAttribute(attribute.LocalName))) {
                return false;
            }

            return TryParsePassword(actual.WorkbookPassword?.Value, out ushort workbookPassword)
                && TryParsePassword(actual.RevisionsPassword?.Value, out ushort revisionsPassword)
                && workbookPassword == expected.WorkbookPassword
                && revisionsPassword == expected.RevisionsPassword
                && (actual.LockStructure?.Value ?? false) == expected.LockStructure
                && (actual.LockWindows?.Value ?? false) == expected.LockWindows
                && (actual.LockRevision?.Value ?? false) == expected.LockRevision;
        }

        internal static bool TryParsePassword(string? value, out ushort password) {
            password = 0;
            return string.IsNullOrWhiteSpace(value)
                || (value!.Length <= 4 && ushort.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out password));
        }

        private static bool IsSupportedAttribute(string localName) =>
            localName == "workbookPassword"
            || localName == "revisionsPassword"
            || localName == "lockStructure"
            || localName == "lockWindows"
            || localName == "lockRevision";
    }
}
