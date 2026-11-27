using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordRoundTripTheme12AtomForWrite = 0x040E;
        private const ushort RecordRoundTripColorMapping12AtomForWrite = 0x040F;
        private const string DrawingThemeNamespaceForWrite =
            "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly DateTimeOffset DeterministicZipTimestamp =
            new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
        private static readonly UTF8Encoding Utf8WithoutBom = new(false);

        /// <summary>
        /// The round-trip records of a DrawingML theme (or theme override, with <paramref name="isOverride"/>) and a
        /// color map (<c>p:clrMap</c>) or color map override (<c>p:clrMapOvr</c>); either may be none.
        /// </summary>
        private static IReadOnlyList<byte[]> BuildRoundTripThemeRecords(
            OfficeXmlNode theme, OfficeXmlNode colorMap, bool isOverride) {
            var records = new List<byte[]>(2);
            if (!theme.IsNone) {
                records.Add(BuildRecord(version: 0, instance: 0,
                    RecordRoundTripTheme12AtomForWrite,
                    BuildRoundTripThemePackage(SerializeLikeSdk(theme),
                        isOverride)));
            }
            if (!colorMap.IsNone) {
                records.Add(BuildRecord(version: 0, instance: 0,
                    RecordRoundTripColorMapping12AtomForWrite,
                    Utf8WithoutBom.GetBytes(
                        BuildRoundTripColorMappingXml(colorMap))));
            }
            return records;
        }

        // The theme markup as the SDK's OuterXml wrote it (an XmlWriter: "<a:x />" for empty elements, whitespace
        // between elements dropped), so the embedded theme package keeps the same bytes.
        private static string SerializeLikeSdk(OfficeXmlNode element) =>
            XElement.Parse(OuterXmlText(element), LoadOptions.None)
                .ToString(SaveOptions.DisableFormatting);

        private static string OuterXmlText(OfficeXmlNode element) {
            byte[] bytes = element.OuterXml();
            return Utf8WithoutBom.GetString(bytes);
        }

        private static byte[] BuildRoundTripThemePackage(string themeXml,
            bool isOverride) {
            if (string.IsNullOrWhiteSpace(themeXml)) {
                throw new InvalidDataException(
                    "A DrawingML theme cannot be empty when writing binary PowerPoint round-trip data.");
            }
            string partName = isOverride
                ? "theme/theme/themeOverride1.xml"
                : "theme/theme/theme1.xml";
            string contentType = isOverride
                ? "application/vnd.openxmlformats-officedocument.themeOverride+xml"
                : "application/vnd.openxmlformats-officedocument.theme+xml";
            string relationshipType = isOverride
                ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride"
                : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
            string contentTypes =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/theme/theme/themeManager.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.themeManager+xml\"/>"
                + $"<Override PartName=\"/{partName}\" ContentType=\"{contentType}\"/>"
                + "</Types>";
            const string rootRelationships =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"theme/theme/themeManager.xml\"/>"
                + "</Relationships>";
            const string themeManager =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<a:themeManager xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>";
            string themeRelationships =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + $"<Relationship Id=\"rId1\" Type=\"{relationshipType}\" Target=\"{Path.GetFileName(partName)}\"/>"
                + "</Relationships>";

            using var output = new MemoryStream();
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create,
                       leaveOpen: true)) {
                WriteThemePackageEntry(archive, "[Content_Types].xml",
                    contentTypes);
                WriteThemePackageEntry(archive, "_rels/.rels",
                    rootRelationships);
                WriteThemePackageEntry(archive,
                    "theme/theme/themeManager.xml", themeManager);
                WriteThemePackageEntry(archive, partName, themeXml);
                WriteThemePackageEntry(archive,
                    "theme/theme/_rels/themeManager.xml.rels",
                    themeRelationships);
            }
            return output.ToArray();
        }

        private static void WriteThemePackageEntry(ZipArchive archive,
            string name, string content) {
            ZipArchiveEntry entry = archive.CreateEntry(name,
                CompressionLevel.Optimal);
            entry.LastWriteTime = DeterministicZipTimestamp;
            using Stream stream = entry.Open();
            byte[] bytes = Utf8WithoutBom.GetBytes(content);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static string BuildRoundTripColorMappingXml(
            OfficeXmlNode colorMap) {
            bool isOverride = string.Equals(colorMap.LocalName, "clrMapOvr",
                StringComparison.Ordinal);
            XElement source = XElement.Parse(OuterXmlText(colorMap),
                LoadOptions.None);
            var target = new XElement(
                XName.Get(isOverride ? "clrMapOvr" : "clrMap",
                    DrawingThemeNamespaceForWrite),
                new XAttribute(XNamespace.Xmlns + "a",
                    DrawingThemeNamespaceForWrite),
                source.Attributes().Where(attribute =>
                        !attribute.IsNamespaceDeclaration)
                    .Select(attribute => new XAttribute(
                        attribute.Name.LocalName, attribute.Value)));
            if (isOverride) target.Add(source.Nodes());
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
                + target.ToString(SaveOptions.DisableFormatting);
        }

        private static bool IsRoundTripThemeRecord(ushort recordType) =>
            recordType is RecordRoundTripTheme12AtomForWrite
                or RecordRoundTripColorMapping12AtomForWrite;
    }
}
