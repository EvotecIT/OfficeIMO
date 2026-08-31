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

        internal static bool TryFindWorkbookPart(
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

                string? target = ReadOfficeDocumentTarget(
                    relationshipsEntry,
                    maxRootRelationshipsBytes);
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
            XDocument document = LoadBoundedXml(stream, maxContentTypesBytes);
            if (document.Root?.Name != XName.Get("Types", PackageContentTypesNamespace)) return false;
            string expectedPartName = "/" + workbookPartName.TrimStart('/');
            string?[] overrides = document.Root
                .Elements(XName.Get("Override", PackageContentTypesNamespace))
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
                string?[] defaults = document.Root
                    .Elements(XName.Get("Default", PackageContentTypesNamespace))
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
            if (string.IsNullOrWhiteSpace(partName) || partName!.IndexOf('\\') >= 0) {
                return string.Empty;
            }

            return "/" + partName.TrimStart('/');
        }

        private static string? ReadOfficeDocumentTarget(
            ZipArchiveEntry relationshipsEntry,
            long maximumBytes) {
            using Stream stream = relationshipsEntry.Open();
            XDocument document = LoadBoundedXml(stream, maximumBytes);
            if (document.Root?.Name != XName.Get("Relationships", PackageRelationshipsNamespace)) return null;
            XElement[] relationships = document.Root
                .Elements(XName.Get("Relationship", PackageRelationshipsNamespace))
                .Where(element =>
                    string.Equals((string?)element.Attribute("Type"), OfficeDocumentRelationship, StringComparison.Ordinal)
                    && !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
                .Take(2)
                .ToArray();
            return relationships.Length == 1 ? (string?)relationships[0].Attribute("Target") : null;
        }

        internal static XDocument LoadBoundedXml(Stream stream, long maximumBytes) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));

            using var bounded = new BootstrapMetadataReadStream(stream, maximumBytes);
            using XmlReader reader = XmlReader.Create(bounded, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CloseInput = false,
                MaxCharactersInDocument = maximumBytes
            });
            return XDocument.Load(reader, LoadOptions.None);
        }

        private static string? NormalizePackageTarget(string target) {
            if (target.IndexOf('\\') >= 0) return null;
            if (ContainsMalformedPercentEscape(target)) return null;
            if (ContainsEncodedPathSeparator(target)) return null;
            string normalized = target;
            if (Uri.TryCreate(normalized, UriKind.Absolute, out _)) return null;
            var packageRoot = new Uri("http://package/", UriKind.Absolute);
            Uri resolved = new Uri(packageRoot, normalized);
            if (!string.Equals(resolved.Host, "package", StringComparison.OrdinalIgnoreCase) ||
                resolved.Query.Length != 0 || resolved.Fragment.Length != 0) return null;
            return Uri.UnescapeDataString(resolved.AbsolutePath).TrimStart('/');
        }

        private static bool ContainsMalformedPercentEscape(string value) {
            for (int index = 0; index < value.Length; index++) {
                if (value[index] != '%') continue;
                if (index > value.Length - 3 || !IsHex(value[index + 1]) || !IsHex(value[index + 2])) return true;
                index += 2;
            }
            return false;
        }

        private static bool IsHex(char value) =>
            value >= '0' && value <= '9' ||
            value >= 'a' && value <= 'f' ||
            value >= 'A' && value <= 'F';

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
                entry.FullName.IndexOf('\\') < 0 &&
                string.Equals(entry.FullName, fullName, StringComparison.OrdinalIgnoreCase));

        private sealed class BootstrapMetadataReadStream : Stream {
            private readonly Stream _inner;
            private readonly long _maximumBytes;
            private long _bytesRead;

            internal BootstrapMetadataReadStream(Stream inner, long maximumBytes) {
                _inner = inner;
                _maximumBytes = maximumBytes;
            }

            public override bool CanRead => _inner.CanRead;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override void Flush() { }

            public override int Read(byte[] buffer, int offset, int count) {
                int read = _inner.Read(buffer, offset, LimitReadSize(count));
                Debit(read);
                return read;
            }

            public override int ReadByte() {
                int value = _inner.ReadByte();
                if (value >= 0) Debit(1);
                return value;
            }

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            private int LimitReadSize(int requested) {
                long remaining = _maximumBytes - _bytesRead;
                return remaining >= requested
                    ? requested
                    : checked((int)Math.Min((long)requested, remaining + 1L));
            }

            private void Debit(int count) {
                if (count < 0 || _bytesRead > _maximumBytes - count) {
                    throw new InvalidDataException(
                        $"The XLSB bootstrap metadata part exceeds the configured limit of {_maximumBytes} bytes.");
                }
                _bytesRead += count;
            }
        }
    }
}
