using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares worksheet markup decoded from native feature records.</summary>
    internal static class XlsbWorksheetFeatureProjector {
        internal static void Apply(ExcelSheet sheet, XlsbSheet source) {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(source);
            if (source.FeatureElements.Count == 0) return;
            OfficeXmlNode worksheet = XlsbWorksheetMarkup.Root(sheet);
            foreach (OfficeXmlNode element in source.ReadFeatureMarkup()) worksheet.Append(element.Clone());
        }

        /// <summary>Compares direct worksheet children against the decoded source markup in record order.</summary>
        internal static bool Matches(OfficeXmlNode worksheet, XlsbSheet source, string name) {
            const int spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
            var actual = worksheet.Children(spreadsheet, name).GetEnumerator();
            foreach (OfficeXmlNode expected in source.ReadFeatureMarkup()) {
                if (!expected.Is(spreadsheet, name)) continue;
                if (!actual.MoveNext() || !actual.Current.OuterXml().AsSpan().SequenceEqual(expected.OuterXml())) return false;
            }
            return !actual.MoveNext();
        }
    }
}
