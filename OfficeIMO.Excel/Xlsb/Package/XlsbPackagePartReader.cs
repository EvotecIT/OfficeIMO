using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel.Xlsb.Package {
    /// <summary>Reads bounded package parts and OPC relationships from an XLSB archive.</summary>
    internal sealed class XlsbPackagePartReader {
        private readonly XlsbImportOptions _options;
        private readonly Dictionary<string, ZipArchiveEntry> _entries;
        private long _decompressedBytesRead;

        internal XlsbPackagePartReader(ZipArchive archive, XlsbImportOptions options) {
            if (archive == null) throw new ArgumentNullException(nameof(archive));
            _options = options ?? throw new ArgumentNullException(nameof(options));
            ValidateDeclaredSizeBudget(archive);
            _entries = BuildEntryIndex(archive);
        }

        internal bool ContainsPart(string partName) => _entries.ContainsKey(NormalizePartName(partName));

        internal byte[] ReadPart(string partName) {
            string normalized = NormalizePartName(partName);
            if (!_entries.TryGetValue(normalized, out ZipArchiveEntry? entry)) {
                throw new InvalidDataException($"The XLSB package part '{normalized}' is missing.");
            }

            if (entry.Length > _options.MaxPartBytes) {
                throw new InvalidDataException($"The XLSB package part '{normalized}' declares {entry.Length} decompressed bytes, exceeding the configured limit of {_options.MaxPartBytes} bytes.");
            }

            int capacity = checked((int)entry.Length);
            using Stream input = entry.Open();
            using var output = new MemoryStream(capacity);
            byte[] buffer = new byte[81920];
            while (true) {
                int read = input.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                if (output.Length + read > _options.MaxPartBytes) {
                    throw new InvalidDataException($"The XLSB package part '{normalized}' exceeds the configured decompression limit of {_options.MaxPartBytes} bytes.");
                }

                _decompressedBytesRead = checked(_decompressedBytesRead + read);
                if (_decompressedBytesRead > _options.MaxPackageBytes) {
                    throw new InvalidDataException($"The XLSB package exceeds the configured aggregate decompression budget of {_options.MaxPackageBytes} bytes.");
                }

                output.Write(buffer, 0, read);
            }

            return output.ToArray();
        }

        internal IReadOnlyDictionary<string, XlsbPackageRelationship> ReadRelationships(string sourcePartName) {
            string relationshipPart = GetRelationshipPartName(sourcePartName);
            if (!ContainsPart(relationshipPart)) {
                return new Dictionary<string, XlsbPackageRelationship>(StringComparer.Ordinal);
            }

            byte[] xml = ReadPart(relationshipPart);
            using var stream = new MemoryStream(xml, writable: false);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CloseInput = false,
                MaxCharactersInDocument = _options.MaxPartBytes
            });
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            var relationships = new Dictionary<string, XlsbPackageRelationship>(StringComparer.Ordinal);
            foreach (XElement element in document.Descendants().Where(item => item.Name.LocalName == "Relationship")) {
                string? id = (string?)element.Attribute("Id");
                string? type = (string?)element.Attribute("Type");
                string? target = (string?)element.Attribute("Target");
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(type) || string.IsNullOrWhiteSpace(target)) {
                    throw new InvalidDataException($"The XLSB relationship part '{relationshipPart}' contains a relationship without Id, Type, or Target.");
                }

                if (relationships.ContainsKey(id!)) {
                    throw new InvalidDataException($"The XLSB relationship part '{relationshipPart}' contains duplicate relationship id '{id}'.");
                }

                relationships.Add(id!, new XlsbPackageRelationship(
                    id!,
                    type!,
                    target!,
                    string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)));
            }

            return relationships;
        }

        internal static string ResolveTarget(string sourcePartName, string target) {
            if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("Relationship target cannot be empty.", nameof(target));
            string normalizedTarget = target.Replace('\\', '/');
            string combined;
            if (normalizedTarget.StartsWith("/", StringComparison.Ordinal)) {
                combined = normalizedTarget.TrimStart('/');
            } else {
                string source = NormalizePartName(sourcePartName);
                int separator = source.LastIndexOf('/');
                string directory = separator < 0 ? string.Empty : source.Substring(0, separator + 1);
                combined = directory + normalizedTarget;
            }

            var segments = new List<string>();
            foreach (string segment in combined.Split('/')) {
                if (segment.Length == 0 || segment == ".") continue;
                if (segment == "..") {
                    if (segments.Count == 0) {
                        throw new InvalidDataException($"The package relationship target '{target}' escapes the package root.");
                    }

                    segments.RemoveAt(segments.Count - 1);
                    continue;
                }

                segments.Add(segment);
            }

            if (segments.Count == 0) {
                throw new InvalidDataException($"The package relationship target '{target}' does not identify a part.");
            }

            return string.Join("/", segments);
        }

        private static Dictionary<string, ZipArchiveEntry> BuildEntryIndex(ZipArchive archive) {
            var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (ZipArchiveEntry entry in archive.Entries) {
                if (string.IsNullOrEmpty(entry.Name)) continue;
                string normalized = NormalizePartName(entry.FullName);
                if (entries.ContainsKey(normalized)) {
                    throw new InvalidDataException($"The XLSB package contains duplicate part name '{normalized}'.");
                }

                entries.Add(normalized, entry);
            }

            return entries;
        }

        private void ValidateDeclaredSizeBudget(ZipArchive archive) {
            long total = 0;
            foreach (ZipArchiveEntry entry in archive.Entries) {
                if (entry.Length > _options.MaxPartBytes) {
                    throw new InvalidDataException($"The XLSB package part '{entry.FullName}' declares {entry.Length} decompressed bytes, exceeding the configured limit of {_options.MaxPartBytes} bytes.");
                }

                try {
                    total = checked(total + entry.Length);
                } catch (OverflowException exception) {
                    throw new InvalidDataException("The XLSB package declares an invalid aggregate decompressed size.", exception);
                }

                if (total > _options.MaxPackageBytes) {
                    throw new InvalidDataException($"The XLSB package declares {total} aggregate decompressed bytes, exceeding the configured limit of {_options.MaxPackageBytes} bytes.");
                }
            }
        }

        private static string GetRelationshipPartName(string sourcePartName) {
            string source = NormalizePartName(sourcePartName);
            int separator = source.LastIndexOf('/');
            string directory = separator < 0 ? string.Empty : source.Substring(0, separator + 1);
            string fileName = separator < 0 ? source : source.Substring(separator + 1);
            return directory + "_rels/" + fileName + ".rels";
        }

        private static string NormalizePartName(string partName) {
            if (string.IsNullOrWhiteSpace(partName)) throw new ArgumentException("Package part name cannot be empty.", nameof(partName));
            string normalized = partName.Replace('\\', '/').TrimStart('/');
            if (normalized.Split('/').Any(segment => segment == "..")) {
                throw new InvalidDataException($"The package part name '{partName}' is not safe.");
            }

            return normalized;
        }
    }
}
