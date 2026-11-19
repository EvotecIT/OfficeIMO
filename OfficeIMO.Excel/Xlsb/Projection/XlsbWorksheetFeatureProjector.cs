using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares the worksheet elements decoded into <see cref="XlsbSheet.FeatureElements"/>.</summary>
    internal static class XlsbWorksheetFeatureProjector {
        internal static void Apply(ExcelSheet sheet, XlsbSheet source) {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(source);
            if (source.FeatureElements.Count == 0) return;
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            foreach (OpenXmlElement element in source.FeatureElements) worksheet.Append(element.CloneNode(true));
        }

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
