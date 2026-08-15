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
            return TryFindWorkbookPart(
                packageBytes,
                MaxRootRelationshipsBytes,
                MaxContentTypesBytes,
                out workbookPartName);
        }

        internal static bool TryFindWorkbookPart(
            byte[] packageBytes,
            long maxRootRelationshipsBytes,
            long maxContentTypesBytes,
            out string? workbookPartName) {
            if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
            if (maxRootRelationshipsBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxRootRelationshipsBytes));
            if (maxContentTypesBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxContentTypesBytes));

            try {
                using var packageStream = new MemoryStream(packageBytes, writable: false);
                using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false);
                return TryFindWorkbookPart(
                    archive,
                    maxRootRelationshipsBytes,
                    maxContentTypesBytes,
                    out workbookPartName);
            } catch (InvalidDataException) {
                workbookPartName = null;
                return false;
            } catch (XmlException) {
                workbookPartName = null;
                return false;
            }
        }

        internal static bool TryFindWorkbookPart(ZipArchive archive, out string? workbookPartName) {
            return TryFindWorkbookPart(
                archive,
                MaxRootRelationshipsBytes,
                MaxContentTypesBytes,
                out workbookPartName);
        }

        private static bool TryFindWorkbookPart(
            ZipArchive archive,
            long maxRootRelationshipsBytes,
            long maxContentTypesBytes,
            out string? workbookPartName) {
            if (archive == null) throw new ArgumentNullException(nameof(archive));

            workbookPartName = null;
            try {
                ZipArchiveEntry? relationshipsEntry = FindEntry(archive, "_rels/.rels");
                if (relationshipsEntry == null || relationshipsEntry.Length > maxRootRelationshipsBytes) {
                    return false;
                }

                string? target = ReadOfficeDocumentTarget(relationshipsEntry);
                if (string.IsNullOrWhiteSpace(target)) {
                    return false;
                }

                string? normalizedTarget = NormalizePackageTarget(target!);
                if (string.IsNullOrWhiteSpace(normalizedTarget) ||
                    !normalizedTarget!.EndsWith(".bin", StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }

                ZipArchiveEntry[] workbookEntries = FindEntries(archive, normalizedTarget).Take(2).ToArray();
                if (workbookEntries.Length != 1 || !HasExcelBinaryWorkbookContentType(
                    archive, normalizedTarget, maxContentTypesBytes)) {
                    return false;
                }

                workbookPartName = workbookEntries[0].FullName;
                return true;
            } catch (InvalidDataException) {
                workbookPartName = null;
                return false;
            } catch (XmlException) {
                workbookPartName = null;
                return false;
            }
        }

        private static bool HasExcelBinaryWorkbookContentType(
            ZipArchive archive,
            string workbookPartName,
            long maxContentTypesBytes) {
            ZipArchiveEntry? contentTypesEntry = FindEntry(archive, "[Content_Types].xml");
            if (contentTypesEntry == null || contentTypesEntry.Length > maxContentTypesBytes) {
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
            string?[] overrides = document
                .Descendants(XName.Get("Override", PackageContentTypesNamespace))
                .Where(element => string.Equals(
                    NormalizeContentTypePartName((string?)element.Attribute("PartName")),
                    expectedPartName,
                    StringComparison.OrdinalIgnoreCase))
                .Select(element => (string?)element.Attribute("ContentType"))
                .Take(2)
                .ToArray();
            if (overrides.Length > 1) return false;
            string? contentType = overrides.SingleOrDefault();

            if (string.IsNullOrWhiteSpace(contentType)) {
                string extension = Path.GetExtension(workbookPartName).TrimStart('.');
                string?[] defaults = document
                    .Descendants(XName.Get("Default", PackageContentTypesNamespace))
                    .Where(element => string.Equals(
                        (string?)element.Attribute("Extension"),
                        extension,
                        StringComparison.OrdinalIgnoreCase))
                    .Select(element => (string?)element.Attribute("ContentType"))
                    .Take(2)
                    .ToArray();
                if (defaults.Length > 1) return false;
                contentType = defaults.SingleOrDefault();
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
            XElement[] relationships = document
                .Descendants(XName.Get("Relationship", PackageRelationshipsNamespace))
                .Where(element =>
                    string.Equals((string?)element.Attribute("Type"), OfficeDocumentRelationship, StringComparison.Ordinal)
                    && !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
                .Take(2)
                .ToArray();
            return relationships.Length == 1 ? (string?)relationships[0].Attribute("Target") : null;
        }

        private static string? NormalizePackageTarget(string target) {
            if (target.IndexOf('\\') >= 0) return null;
            if (ContainsEncodedPathSeparator(target)) return null;
            string normalized = target;
            if (Uri.TryCreate(normalized, UriKind.Absolute, out _)) return null;
            var packageRoot = new Uri("http://package/", UriKind.Absolute);
            Uri resolved = new Uri(packageRoot, normalized);
            if (!string.Equals(resolved.Host, "package", StringComparison.OrdinalIgnoreCase) ||
                resolved.Query.Length != 0 || resolved.Fragment.Length != 0) return null;
            return Uri.UnescapeDataString(resolved.AbsolutePath).TrimStart('/');
        }

        private static bool ContainsEncodedPathSeparator(string value) {
            for (int index = 0; index <= value.Length - 3; index++) {
                if (value[index] != '%') continue;
                char high = char.ToLowerInvariant(value[index + 1]);
                char low = char.ToLowerInvariant(value[index + 2]);
                if (high == '2' && low == 'f' || high == '5' && low == 'c') return true;
            }
            return false;
        }

        private static ZipArchiveEntry? FindEntry(ZipArchive archive, string fullName) {
            return FindEntries(archive, fullName).FirstOrDefault();
        }

        private static IEnumerable<ZipArchiveEntry> FindEntries(ZipArchive archive, string fullName) =>
            archive.Entries.Where(entry =>
                string.Equals(entry.FullName.Replace('\\', '/'), fullName, StringComparison.OrdinalIgnoreCase));
    }
}
