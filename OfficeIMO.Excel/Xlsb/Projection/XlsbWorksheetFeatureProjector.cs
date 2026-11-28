using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
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

        // The existing table path still uses the typed comparison.
        /// <summary>True when the worksheet's <typeparamref name="T"/> children equal the ones decoded from the source.</summary>
        internal static bool Matches<T>(ExcelSheet sheet, XlsbSheet source) where T : OpenXmlElement {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            string[] actual = worksheet.Elements<T>().Select(element => element.OuterXml).ToArray();
            string[] expected = source.FeatureElements.OfType<T>().Select(element => element.OuterXml).ToArray();
            return actual.SequenceEqual(expected, StringComparer.Ordinal);
        }
    }
}
