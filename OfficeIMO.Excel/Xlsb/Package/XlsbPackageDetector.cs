using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel.Xlsb.Package {
    /// <summary>
    /// Identifies the workbook binary part through the package-level Office document relationship.
    /// </summary>
    internal static class XlsbPackageDetector {
        private const string OfficeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const int MaxRootRelationshipsBytes = 1024 * 1024;
        private const int MaxContentTypesBytes = 1024 * 1024;

        internal static bool TryFindWorkbookPart(byte[] packageBytes, out string? workbookPartName) {
            if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));

            workbookPartName = null;
            try {
                using var packageStream = new MemoryStream(packageBytes, writable: false);
                using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false);
                ZipArchiveEntry? relationshipsEntry = FindEntry(archive, "_rels/.rels");
                if (relationshipsEntry == null || relationshipsEntry.Length > MaxRootRelationshipsBytes) {
                    return false;
                }

                string? target = ReadOfficeDocumentTarget(relationshipsEntry);
                if (string.IsNullOrWhiteSpace(target)) {
                    return false;
                }

                string normalizedTarget = NormalizePackageTarget(target!);
                if (!normalizedTarget.EndsWith(".bin", StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }

                ZipArchiveEntry? workbookEntry = FindEntry(archive, normalizedTarget);
                if (workbookEntry == null || !HasExcelBinaryWorkbookContentType(archive, normalizedTarget)) {
                    return false;
                }

                workbookPartName = workbookEntry.FullName;
                return true;
            } catch (InvalidDataException) {
                return false;
            } catch (XmlException) {
                return false;
            }
        }

        private static bool HasExcelBinaryWorkbookContentType(ZipArchive archive, string workbookPartName) {
            ZipArchiveEntry? contentTypesEntry = FindEntry(archive, "[Content_Types].xml");
            if (contentTypesEntry == null || contentTypesEntry.Length > MaxContentTypesBytes) {
                return false;
            }

            using Stream stream = contentTypesEntry.Open();
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CloseInput = false
            });
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            string expectedPartName = "/" + workbookPartName.TrimStart('/');
            string? contentType = document
                .Descendants()
                .Where(element => element.Name.LocalName == "Override")
                .Where(element => string.Equals(
                    NormalizeContentTypePartName((string?)element.Attribute("PartName")),
                    expectedPartName,
                    StringComparison.OrdinalIgnoreCase))
                .Select(element => (string?)element.Attribute("ContentType"))
                .FirstOrDefault();

            return !string.IsNullOrWhiteSpace(contentType)
                && contentType!.StartsWith("application/vnd.ms-excel", StringComparison.OrdinalIgnoreCase)
                && (contentType.IndexOf("binary", StringComparison.OrdinalIgnoreCase) >= 0
                    || contentType.EndsWith(".main", StringComparison.OrdinalIgnoreCase));
        }

        private static string NormalizeContentTypePartName(string? partName) {
            if (string.IsNullOrWhiteSpace(partName)) {
                return string.Empty;
            }

            return "/" + partName!.Replace('\\', '/').TrimStart('/');
        }

        private static string? ReadOfficeDocumentTarget(ZipArchiveEntry relationshipsEntry) {
            using Stream stream = relationshipsEntry.Open();
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CloseInput = false
            });
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            XElement? relationship = document
                .Descendants()
                .FirstOrDefault(element =>
                    element.Name.LocalName == "Relationship"
                    && string.Equals((string?)element.Attribute("Type"), OfficeDocumentRelationship, StringComparison.Ordinal)
                    && !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase));
            return (string?)relationship?.Attribute("Target");
        }

        private static string NormalizePackageTarget(string target) {
            string normalized = target.Replace('\\', '/').TrimStart('/');
            while (normalized.StartsWith("./", StringComparison.Ordinal)) {
                normalized = normalized.Substring(2);
            }

            return normalized;
        }

        private static ZipArchiveEntry? FindEntry(ZipArchive archive, string fullName) {
            return archive.Entries.FirstOrDefault(entry =>
                string.Equals(entry.FullName.Replace('\\', '/'), fullName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
