using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private static readonly XNamespace CoreNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private static readonly XNamespace DublinCoreNamespace = "http://purl.org/dc/elements/1.1/";
        private static readonly XNamespace DublinTermsNamespace = "http://purl.org/dc/terms/";
        private static readonly XNamespace SchemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance";

        /// <summary>Copies a loaded package and replaces only its core properties when nothing else changed.</summary>
        private bool TryWritePropertiesOnlyPackage(Stream destination, ExcelSaveOptions? options) {
            byte[]? source = _propertiesOnlySourceBytes;
            if (source == null
                || options?.DisableFastPackageWriter == true
                || options?.SafePreflight == true
                || options?.SafeRepairDefinedNames == true
                || options?.ValidateOpenXml == true
                || HasCalculationSaveWork(options)
                || _spreadSheetDocument.DigitalSignatureOriginPart != null
                || _spreadSheetDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null) {
                return false;
            }

            _spreadSheetDocument.PackageProperties.Modified = DateTime.UtcNow;
            byte[]? package = ExcelPackageEntrySplicer.TryReplace(source, RewriteCoreProperties);
            if (package == null) return false;

            ThrowIfPackageMaterializationExceedsLimit(package.LongLength, options);
            PrepareDestinationStreamForWrite(destination);
            destination.Write(package, 0, package.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            MarkPackageClean(package);
            return true;
        }

        private KeyValuePair<string, byte[]> RewriteCoreProperties(ExcelPackageEntrySplicer.EntryReader read) {
            string entryName = FindCorePropertiesEntry(read("_rels/.rels"));
            XDocument document = LoadCoreXml(read(entryName));
            XElement root = document.Root ?? throw new InvalidDataException("Core properties are empty.");
            var properties = _spreadSheetDocument.PackageProperties;
            SetCoreText(root, DublinCoreNamespace + "title", properties.Title);
            SetCoreText(root, DublinCoreNamespace + "creator", properties.Creator);
            SetCoreText(root, DublinCoreNamespace + "subject", properties.Subject);
            SetCoreText(root, CoreNamespace + "category", properties.Category);
            SetCoreText(root, DublinCoreNamespace + "description", properties.Description);
            SetCoreText(root, CoreNamespace + "keywords", properties.Keywords);
            SetCoreText(root, CoreNamespace + "lastModifiedBy", properties.LastModifiedBy);
            SetCoreText(root, CoreNamespace + "version", properties.Version);
            SetCoreDate(root, DublinTermsNamespace + "created", properties.Created, typed: true);
            SetCoreDate(root, DublinTermsNamespace + "modified", properties.Modified, typed: true);
            SetCoreDate(root, CoreNamespace + "lastPrinted", properties.LastPrinted, typed: false);

            using var output = new MemoryStream();
            using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false) })) {
                document.Save(writer);
            }
            return new KeyValuePair<string, byte[]>(entryName, output.ToArray());
        }

        private static XDocument LoadCoreXml(byte[] xml) {
            using var input = new MemoryStream(xml, writable: false);
            using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
            return XDocument.Load(reader);
        }

        private static string FindCorePropertiesEntry(byte[] relationships) {
            XNamespace package = "http://schemas.openxmlformats.org/package/2006/relationships";
            XElement relationship = LoadCoreXml(relationships).Root?.Elements(package + "Relationship")
                .SingleOrDefault(candidate => string.Equals((string?)candidate.Attribute("Type"), CorePropertiesRelationshipType, StringComparison.Ordinal))
                ?? throw new InvalidDataException("The package has no core properties.");
            return ((string?)relationship.Attribute("Target"))?.TrimStart('/')
                ?? throw new InvalidDataException("The core properties relationship has no target.");
        }

        private static void SetCoreText(XElement root, XName name, string? value) {
            XElement? element = root.Element(name);
            if (value == null) {
                element?.Remove();
            } else if (element == null) {
                root.Add(new XElement(name, value));
            } else {
                element.Value = value;
            }
        }

        private static void SetCoreDate(XElement root, XName name, DateTime? value, bool typed) {
            if (value == null) {
                root.Element(name)?.Remove();
                return;
            }

            DateTime utc = value.Value.Kind == DateTimeKind.Local ? value.Value.ToUniversalTime() : value.Value;
            SetCoreText(root, name, utc.ToString("yyyy-MM-ddTHH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture));
            if (typed) {
                root.Element(name)!.SetAttributeValue(SchemaInstanceNamespace + "type", "dcterms:W3CDTF");
            }
        }
    }

    /// <summary>Rewrites one ZIP package with selected entries replaced and every other entry copied byte for byte.</summary>
    internal static class ExcelPackageEntrySplicer {
        internal delegate byte[] EntryReader(string name);

        private const uint LocalHeaderSignature = 0x04034B50;
        private const uint CentralHeaderSignature = 0x02014B50;
        private const uint EndOfCentralDirectorySignature = 0x06054B50;
        private const int LocalHeaderSize = 30;
        private const int CentralHeaderSize = 46;
        private const int EndRecordSize = 22;

        /// <summary>
        /// Returns the package with the entry <paramref name="replace"/> names rewritten,
        /// or null when the package is outside the plain ZIP shape handled here.
        /// </summary>
        internal static byte[]? TryReplace(byte[] package, Func<EntryReader, KeyValuePair<string, byte[]>> replace) {
            try {
                return Replace(package, replace);
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or ArgumentException or InvalidOperationException or IndexOutOfRangeException) {
                return null;
            }
        }

        private static byte[]? Replace(byte[] package, Func<EntryReader, KeyValuePair<string, byte[]>> replace) {
            int end = FindEndRecord(package);
            if (end < 0) return null;
            int count = ReadUInt16(package, end + 10);
            long directorySize = ReadUInt32(package, end + 12);
            long directoryOffset = ReadUInt32(package, end + 16);
            if (count == 0xFFFF || directoryOffset == 0xFFFFFFFFL || ReadUInt16(package, end + 4) != 0) return null;

            var entries = new List<Entry>(count);
            int cursor = (int)directoryOffset;
            for (int index = 0; index < count; index++) {
                if (ReadUInt32(package, cursor) != CentralHeaderSignature) return null;
                int nameLength = ReadUInt16(package, cursor + 28);
                int headerLength = CentralHeaderSize + nameLength + ReadUInt16(package, cursor + 30) + ReadUInt16(package, cursor + 32);
                long localOffset = ReadUInt32(package, cursor + 42);
                if (localOffset == 0xFFFFFFFFL) return null;
                entries.Add(new Entry(System.Text.Encoding.UTF8.GetString(package, cursor + CentralHeaderSize, nameLength), cursor, headerLength, (int)localOffset));
                cursor += headerLength;
            }

            var ordered = entries.OrderBy(static entry => entry.LocalOffset).ToArray();
            for (int index = 0; index < ordered.Length; index++) {
                ordered[index].LocalEnd = index + 1 < ordered.Length ? ordered[index + 1].LocalOffset : (int)directoryOffset;
            }

            byte[] Read(string name) => ReadEntry(package, entries.Find(entry => string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidDataException($"Package entry '{name}' is missing."));

            KeyValuePair<string, byte[]> replacement = replace(Read);
            bool replaced = false;
            using var output = new MemoryStream(package.Length + 512);
            using var directory = new MemoryStream((int)directorySize + 64);
            foreach (Entry entry in ordered) {
                int localOffset = (int)output.Position;
                if (!string.Equals(entry.Name, replacement.Key, StringComparison.OrdinalIgnoreCase)) {
                    output.Write(package, entry.LocalOffset, entry.LocalEnd - entry.LocalOffset);
                    directory.Write(package, entry.CentralOffset, entry.CentralLength);
                    WriteUInt32(directory, directory.Position - entry.CentralLength + 42, (uint)localOffset);
                    continue;
                }

                replaced = true;
                WriteReplacement(package, entry, replacement.Value, output, directory, localOffset);
            }
            if (!replaced) return null;

            int directoryStart = (int)output.Position;
            directory.Position = 0;
            directory.CopyTo(output);
            int endStart = (int)output.Position;
            output.Write(package, end, package.Length - end);
            WriteUInt32(output, endStart + 12, (uint)(endStart - directoryStart));
            WriteUInt32(output, endStart + 16, (uint)directoryStart);
            return output.ToArray();
        }

        private static void WriteReplacement(byte[] package, Entry entry, byte[] content, MemoryStream output, MemoryStream directory, int localOffset) {
            byte[] compressed;
            using (var deflated = new MemoryStream(content.Length / 2 + 64)) {
                using (var deflate = new DeflateStream(deflated, CompressionLevel.Optimal, leaveOpen: true)) {
                    deflate.Write(content, 0, content.Length);
                }
                compressed = deflated.ToArray();
            }

            int nameLength = ReadUInt16(package, entry.CentralOffset + 28);
            uint crc = OfficePngCrc32.Complete(OfficePngCrc32.Append(OfficePngCrc32.Begin(), content, 0, content.Length));
            var local = new byte[LocalHeaderSize];
            WriteUInt32(local, 0, LocalHeaderSignature);
            Buffer.BlockCopy(package, entry.CentralOffset + 6, local, 4, 2);
            WriteUInt16(local, 6, (ushort)(ReadUInt16(package, entry.CentralOffset + 8) & ~0x0008));
            WriteUInt16(local, 8, 8);
            Buffer.BlockCopy(package, entry.CentralOffset + 12, local, 10, 4);
            WriteUInt32(local, 14, crc);
            WriteUInt32(local, 18, (uint)compressed.Length);
            WriteUInt32(local, 22, (uint)content.Length);
            WriteUInt16(local, 26, (ushort)nameLength);
            output.Write(local, 0, local.Length);
            output.Write(package, entry.CentralOffset + CentralHeaderSize, nameLength);
            output.Write(compressed, 0, compressed.Length);

            var central = new byte[entry.CentralLength];
            Buffer.BlockCopy(package, entry.CentralOffset, central, 0, central.Length);
            WriteUInt16(central, 8, (ushort)(ReadUInt16(central, 8) & ~0x0008));
            WriteUInt16(central, 10, 8);
            WriteUInt32(central, 16, crc);
            WriteUInt32(central, 20, (uint)compressed.Length);
            WriteUInt32(central, 24, (uint)content.Length);
            WriteUInt32(central, 42, (uint)localOffset);
            directory.Write(central, 0, central.Length);
        }

        private static byte[] ReadEntry(byte[] package, Entry entry) {
            int nameLength = ReadUInt16(package, entry.LocalOffset + 26);
            int extraLength = ReadUInt16(package, entry.LocalOffset + 28);
            if (ReadUInt32(package, entry.LocalOffset) != LocalHeaderSignature) throw new InvalidDataException("Invalid local header.");
            int method = ReadUInt16(package, entry.CentralOffset + 10);
            int compressedSize = checked((int)ReadUInt32(package, entry.CentralOffset + 20));
            int uncompressedSize = checked((int)ReadUInt32(package, entry.CentralOffset + 24));
            int data = entry.LocalOffset + LocalHeaderSize + nameLength + extraLength;
            if (method == 0) {
                var stored = new byte[uncompressedSize];
                Buffer.BlockCopy(package, data, stored, 0, stored.Length);
                return stored;
            }
            if (method != 8) throw new InvalidDataException("Unsupported compression.");
            using var compressed = new MemoryStream(package, data, compressedSize, writable: false);
            using var deflate = new DeflateStream(compressed, CompressionMode.Decompress);
            var content = new byte[uncompressedSize];
            int read = 0;
            while (read < content.Length) {
                int block = deflate.Read(content, read, content.Length - read);
                if (block == 0) throw new InvalidDataException("Truncated package entry.");
                read += block;
            }
            return content;
        }

        private static int FindEndRecord(byte[] package) {
            int lowest = Math.Max(0, package.Length - EndRecordSize - ushort.MaxValue);
            for (int index = package.Length - EndRecordSize; index >= lowest; index--) {
                if (ReadUInt32(package, index) == EndOfCentralDirectorySignature
                    && index + EndRecordSize + ReadUInt16(package, index + 20) == package.Length) {
                    return index;
                }
            }
            return -1;
        }

        private static int ReadUInt16(byte[] data, int offset) => data[offset] | data[offset + 1] << 8;

        private static long ReadUInt32(byte[] data, int offset) =>
            (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24);

        private static void WriteUInt16(byte[] data, int offset, ushort value) {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
        }

        private static void WriteUInt32(byte[] data, int offset, uint value) {
            WriteUInt16(data, offset, (ushort)value);
            WriteUInt16(data, offset + 2, (ushort)(value >> 16));
        }

        private static void WriteUInt32(MemoryStream stream, long offset, uint value) {
            long position = stream.Position;
            stream.Position = offset;
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
            stream.Position = position;
        }

        private sealed class Entry {
            internal Entry(string name, int centralOffset, int centralLength, int localOffset) {
                Name = name;
                CentralOffset = centralOffset;
                CentralLength = centralLength;
                LocalOffset = localOffset;
            }

            internal string Name { get; }
            internal int CentralOffset { get; }
            internal int CentralLength { get; }
            internal int LocalOffset { get; }
            internal int LocalEnd { get; set; }
        }
    }
}
