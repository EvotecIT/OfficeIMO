using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel.Xlsb.Package {
    /// <summary>
    /// Identifies the workbook binary part through the package-level Office document relationship.
    /// </summary>
    internal static class XlsbPackageDetector {
        private const string OfficeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string PackageContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const int MaxRootRelationshipsBytes = 1024 * 1024;
        private const int MaxContentTypesBytes = 1024 * 1024;

        internal static bool TryFindWorkbookPart(byte[] packageBytes, out string? workbookPartName) {
            if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));

            try {
                using var packageStream = new MemoryStream(packageBytes, writable: false);
                using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false);
                return TryFindWorkbookPart(archive, out workbookPartName);
            } catch (InvalidDataException) {
                workbookPartName = null;
                return false;
            } catch (XmlException) {
                workbookPartName = null;
                return false;
            }
        }

        internal static bool TryFindWorkbookPart(ZipArchive archive, out string? workbookPartName) {
            if (archive == null) throw new ArgumentNullException(nameof(archive));

            workbookPartName = null;
            try {
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
                workbookPartName = null;
                return false;
            } catch (XmlException) {
                workbookPartName = null;
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
            if (document.Root?.Name != XName.Get("Types", PackageContentTypesNamespace)) return false;
            string expectedPartName = "/" + workbookPartName.TrimStart('/');
            string? contentType = document
                .Descendants(XName.Get("Override", PackageContentTypesNamespace))
                .Where(element => string.Equals(
                    NormalizeContentTypePartName((string?)element.Attribute("PartName")),
                    expectedPartName,
                    StringComparison.OrdinalIgnoreCase))
                .Select(element => (string?)element.Attribute("ContentType"))
                .FirstOrDefault();

            if (string.IsNullOrWhiteSpace(contentType)) {
                string extension = Path.GetExtension(workbookPartName).TrimStart('.');
                contentType = document
                    .Descendants(XName.Get("Default", PackageContentTypesNamespace))
                    .Where(element => string.Equals(
                        (string?)element.Attribute("Extension"),
                        extension,
                        StringComparison.OrdinalIgnoreCase))
                    .Select(element => (string?)element.Attribute("ContentType"))
                    .FirstOrDefault();
            }

            return string.Equals(
                contentType,
                "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
                StringComparison.OrdinalIgnoreCase);
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
            if (document.Root?.Name != XName.Get("Relationships", PackageRelationshipsNamespace)) return null;
            XElement? relationship = document
                .Descendants(XName.Get("Relationship", PackageRelationshipsNamespace))
                .FirstOrDefault(element =>
                    string.Equals((string?)element.Attribute("Type"), OfficeDocumentRelationship, StringComparison.Ordinal)
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
