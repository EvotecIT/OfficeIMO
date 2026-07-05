using System.IO.Compression;
using System.Xml.Linq;

namespace OfficeIMO.Excel.LegacyXls.Model {
    internal static class LegacyXlsThemePackageReader {
        private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string ThemeEntryPath = "theme/theme/theme1.xml";

        internal static bool TryExtractThemeXml(LegacyXlsThemeRecord themeRecord, out string? themeXml) {
            if (themeRecord == null) {
                throw new ArgumentNullException(nameof(themeRecord));
            }

            themeXml = null;
            if (!themeRecord.HasThemeBytes) {
                return false;
            }

            try {
                using var stream = new MemoryStream(themeRecord.ThemeBytes);
                using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
                ZipArchiveEntry? entry = archive.GetEntry(ThemeEntryPath);
                if (entry == null) {
                    return false;
                }

                using Stream entryStream = entry.Open();
                using var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                string xml = reader.ReadToEnd();
                XDocument document = XDocument.Parse(xml);
                if (document.Root?.Name != DrawingNamespace + "theme") {
                    return false;
                }

                themeXml = xml;
                return true;
            } catch (InvalidDataException) {
                return false;
            } catch (System.Xml.XmlException) {
                return false;
            }
        }
    }
}
