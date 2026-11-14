using System.Collections.Concurrent;
using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf {

    /// <summary>Offline Adobe predefined CJK CMaps and their character collections' CID-to-Unicode maps.</summary>
    /// <remarks>Bundled data covers the predefined CMaps of Adobe-GB1, Adobe-CNS1, Adobe-Japan1, Adobe-Korea1 and
    /// Adobe-KR. Each collection is one embedded resource; a CMap and a Unicode map are inflated only when a
    /// document first names them. A document-supplied encoding CMap contributes two-byte code-to-CID ranges alone,
    /// parsed within the fixed 65,536-code table.</remarks>
    internal sealed partial class PdfPredefinedCMap {
        private static readonly ConcurrentDictionary<string, Lazy<PdfPredefinedCMap>> Bundled = new(StringComparer.Ordinal);
        private static readonly ConcurrentDictionary<string, Lazy<CollectionIndex>> Collections = new(StringComparer.Ordinal);
        private static readonly ConcurrentDictionary<string, Lazy<CidUnicodeMap>> CollectionUnicode = new(StringComparer.Ordinal);
        private static readonly Regex CodespaceSectionRegex = new(@"begincodespacerange([\s\S]*?)endcodespacerange", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex CidRangeSectionRegex = new(@"begincidrange([\s\S]*?)endcidrange", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex CidCharSectionRegex = new(@"begincidchar([\s\S]*?)endcidchar", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex HexCodeRegex = new(@"<\s*([0-9A-Fa-f]{1,8})\s*>", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex CidRangeEntryRegex = new(@"<\s*([0-9A-Fa-f]{4})\s*>\s*<\s*([0-9A-Fa-f]{4})\s*>\s*(\d{1,5})", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex CidCharEntryRegex = new(@"<\s*([0-9A-Fa-f]{4})\s*>\s*(\d{1,5})", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private const long MaximumEmbeddedCMapAssignments = 1_000_000L;
        private const int MaximumCodeLength = 4;
        private const string UnicodeEntry = "#Unicode";

        /// <summary>A map that decodes nothing, used when an embedded CMap cannot be read.</summary>
        internal static PdfPredefinedCMap Empty { get; } = new(new ushort[65536], new CidUnicodeMap(Array.Empty<int>(), null));

        // Embedded encoding CMaps keep the fixed two-byte table.
        private readonly ushort[]? _twoByteCids;
        // Bundled CMaps: codespace ranges and code-to-CID runs indexed by code length.
        private readonly CodespaceRange[] _codespace;
        private readonly CodeRuns?[] _cidRuns;
        private readonly CodeRuns?[] _notdefRuns;
        private readonly PdfPredefinedCMap? _parent;
        private readonly CidUnicodeMap _unicode;

        private PdfPredefinedCMap(ushort[] cids, CidUnicodeMap unicode) {
            ulong hash = 14695981039346656037UL;
            foreach (ushort cid in cids) hash = (hash ^ cid) * 1099511628211UL;
            Identity = "embedded:" + hash.ToString("x16", CultureInfo.InvariantCulture);
            _twoByteCids = cids;
            _codespace = Array.Empty<CodespaceRange>();
            _cidRuns = Array.Empty<CodeRuns?>();
            _notdefRuns = Array.Empty<CodeRuns?>();
            _unicode = unicode;
        }

        private PdfPredefinedCMap(string name, CodespaceRange[] codespace, CodeRuns?[] cidRuns, CodeRuns?[] notdefRuns,
            PdfPredefinedCMap? parent, CidUnicodeMap unicode) {
            Identity = name;
            _codespace = codespace.Length == 0 && parent != null ? parent._codespace : codespace;
            _cidRuns = cidRuns;
            _notdefRuns = notdefRuns;
            _parent = parent;
            _unicode = unicode;
        }

        /// <summary>Distinguishes code-to-CID maps: the bundled CMap name, or a digest of an embedded CMap's table.</summary>
        internal string Identity { get; }

        /// <summary>Returns a shared lazy map only when the CMap belongs to the descendant's Adobe character collection.</summary>
        internal static Lazy<PdfPredefinedCMap>? Find(string encoding, string registry, string ordering) {
            if (!string.Equals(registry, "Adobe", StringComparison.Ordinal)) return null;
            string? bundledOrdering = GetBundledOrdering(encoding);
            if (bundledOrdering == null || !string.Equals(bundledOrdering, ordering, StringComparison.Ordinal)) return null;
            return Bundled.GetOrAdd(encoding, static name => new Lazy<PdfPredefinedCMap>(() => Load(name)));
        }

        /// <summary>True when the bundled resources map this character collection's CIDs to Unicode.</summary>
        internal static bool HasUnicodeCollection(string registry, string ordering) =>
            string.Equals(registry, "Adobe", StringComparison.Ordinal) && IsBundledOrdering(ordering);

        /// <summary>
        /// Reads the two-byte code-to-CID ranges of an embedded encoding CMap. Returns <see cref="Empty"/> when the
        /// CMap uses other code lengths, a usecmap parent, or malformed ranges, so text keeps the unmapped diagnostic.
        /// </summary>
        internal static PdfPredefinedCMap FromEmbeddedCMap(byte[] data, string ordering) {
            if (!IsBundledOrdering(ordering)) return Empty;
            string text = Encoding.ASCII.GetString(data);
            if (text.Contains("usecmap")) return Empty;
            MatchCollection codespaces = CodespaceSectionRegex.Matches(text);
            if (codespaces.Count == 0) return Empty;
            foreach (Match section in codespaces) {
                foreach (Match code in HexCodeRegex.Matches(section.Groups[1].Value)) {
                    if (code.Groups[1].Value.Length != 4) return Empty;
                }
            }

            var cids = new ushort[65536];
            long assignments = 0L;
            foreach (Match section in CidRangeSectionRegex.Matches(text)) {
                foreach (Match entry in CidRangeEntryRegex.Matches(section.Groups[1].Value)) {
                    int first = Convert.ToInt32(entry.Groups[1].Value, 16);
                    int last = Convert.ToInt32(entry.Groups[2].Value, 16);
                    int cid = int.Parse(entry.Groups[3].Value, CultureInfo.InvariantCulture);
                    if (first > last || (long)cid + (last - first) > ushort.MaxValue) return Empty;
                    assignments += last - first + 1;
                    if (assignments > MaximumEmbeddedCMapAssignments) return Empty;
                    for (int code = first; code <= last; code++) cids[code] = (ushort)(cid + code - first);
                }
            }
            foreach (Match section in CidCharSectionRegex.Matches(text)) {
                foreach (Match entry in CidCharEntryRegex.Matches(section.Groups[1].Value)) {
                    int cid = int.Parse(entry.Groups[2].Value, CultureInfo.InvariantCulture);
                    if (cid > ushort.MaxValue || ++assignments > MaximumEmbeddedCMapAssignments) return Empty;
                    cids[Convert.ToInt32(entry.Groups[1].Value, 16)] = (ushort)cid;
                }
            }
            return new PdfPredefinedCMap(cids, GetCollectionUnicode(ordering));
        }

        /// <summary>
        /// Returns the byte length of the code starting at <paramref name="offset"/>. The result can exceed the remaining
        /// bytes when the string ends inside a code; bytes outside every codespace range use the shortest plausible length.
        /// </summary>
        internal int GetCodeLength(byte[] bytes, int offset) => ReadCode(bytes, offset, out _, out _);

        /// <summary>True when the bytes divide into complete codes of this CMap.</summary>
        internal bool HasCompleteCodes(byte[] bytes, int offset, int count) {
            int end = offset + count;
            for (int index = offset; index < end;) {
                int length = ReadCode(bytes, index, out _, out _);
                if (index + length > end) return false;
                index += length;
            }
            return true;
        }

        /// <summary>Maps one complete, valid code to its CID; undefined codes and notdef CID 0 have no mapping.</summary>
        internal bool TryGetCid(byte[] code, out int cid) {
            cid = 0;
            if (code.Length == 0 || ReadCode(code, 0, out uint value, out bool valid) != code.Length || !valid) return false;
            cid = Lookup(code.Length, value);
            return cid != 0;
        }

        /// <summary>Resolves a valid painted code, including the encoding's default CID 0.</summary>
        internal bool TryGetPaintedCid(byte[] bytes, int offset, int length, out int cid) {
            cid = 0;
            if (offset < 0 || length <= 0 || offset + length > bytes.Length ||
                ReadCode(bytes, offset, out uint value, out bool valid) != length || !valid) return false;
            cid = Lookup(length, value);
            return true;
        }

        /// <summary>Decodes every painted code through both authoritative maps, refusing incomplete coverage.</summary>
        internal bool TryDecode(byte[] bytes, int maximumCharacters, out string decoded) {
            decoded = string.Empty;
            var text = new StringBuilder(bytes.Length);
            for (int index = 0; index < bytes.Length;) {
                int length = ReadCode(bytes, index, out uint value, out bool valid);
                if (!valid || index + length > bytes.Length) return false;
                int cid = Lookup(length, value);
                if (cid == 0 || !_unicode.TryAppend(cid, text)) return false;
                if (text.Length > maximumCharacters) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maximumCharacters, text.Length);
                }
                index += length;
            }
            decoded = text.ToString();
            return true;
        }

        /// <summary>Unicode values of the character collection in CID order; sequence-valued CIDs report scalar -1.</summary>
        internal IEnumerable<(int Cid, int Scalar)> EnumerateCollectionScalars() => _unicode.EnumerateScalars();

        // Reads one code. A complete match of a codespace range is valid. Otherwise the code is invalid and its
        // length is that of a range whose leading bytes match, or the shortest range when none does.
        private int ReadCode(byte[] bytes, int offset, out uint value, out bool valid) {
            value = 0;
            valid = false;
            if (_twoByteCids != null) {
                if (offset + 1 >= bytes.Length) return 2;
                value = (uint)((bytes[offset] << 8) | bytes[offset + 1]);
                valid = true;
                return 2;
            }
            int remaining = bytes.Length - offset;
            int partialLength = 0;
            int firstByteLength = 0;
            int shortest = MaximumCodeLength;
            foreach (CodespaceRange range in _codespace) {
                shortest = Math.Min(shortest, range.Length);
                int matched = range.MatchingPrefix(bytes, offset, Math.Min(range.Length, remaining));
                if (matched == range.Length) {
                    for (int index = 0; index < range.Length; index++) value = (value << 8) | bytes[offset + index];
                    valid = true;
                    return range.Length;
                }
                if (matched == remaining && partialLength == 0) partialLength = range.Length;
                if (matched > 0 && firstByteLength == 0) firstByteLength = range.Length;
            }
            return partialLength > 0 ? partialLength : firstByteLength > 0 ? firstByteLength : shortest;
        }

        private int Lookup(int length, uint code) {
            for (PdfPredefinedCMap? map = this; map != null; map = map._parent) {
                if (map._twoByteCids != null) return length == 2 && code <= ushort.MaxValue ? map._twoByteCids[code] : 0;
                if (length < map._cidRuns.Length && map._cidRuns[length]?.TryGet(code, out int cid) == true) return cid;
            }
            for (PdfPredefinedCMap? map = this; map != null; map = map._parent) {
                if (length < map._notdefRuns.Length && map._notdefRuns[length]?.TryGet(code, out int cid) == true) return cid;
            }
            return 0;
        }

        private static PdfPredefinedCMap Load(string name) {
            string ordering = GetBundledOrdering(name) ?? throw new InvalidDataException("Unknown bundled Adobe CMap: " + name);
            var reader = new PayloadReader(ReadEntry(ordering, name));
            reader.ReadVarint(); // WMode: vertical writing is selected from the CMap name.
            string parentName = reader.ReadText();
            var codespace = new CodespaceRange[reader.ReadVarint()];
            for (int index = 0; index < codespace.Length; index++) {
                int length = reader.ReadCodeLength();
                codespace[index] = new CodespaceRange(length, reader.ReadBytes(length), reader.ReadBytes(length));
            }
            CodeRuns?[] cids = ReadRuns(reader, constant: false);
            CodeRuns?[] notdef = ReadRuns(reader, constant: true);
            if (!reader.AtEnd) throw new InvalidDataException("Invalid embedded Adobe CMap: " + name);
            PdfPredefinedCMap? parent = null;
            if (parentName.Length > 0) {
                if (!string.Equals(GetBundledOrdering(parentName), ordering, StringComparison.Ordinal)) {
                    throw new InvalidDataException("Invalid embedded Adobe CMap parent: " + parentName);
                }
                parent = Bundled.GetOrAdd(parentName, static key => new Lazy<PdfPredefinedCMap>(() => Load(key))).Value;
            }
            if (codespace.Length == 0 && parent == null) throw new InvalidDataException("Embedded Adobe CMap has no codespace: " + name);
            return new PdfPredefinedCMap(name, codespace, cids, notdef, parent, GetCollectionUnicode(ordering));
        }

        private static CodeRuns?[] ReadRuns(PayloadReader reader, bool constant) {
            var runs = new CodeRuns?[MaximumCodeLength + 1];
            int groups = reader.ReadVarint();
            for (int group = 0; group < groups; group++) {
                int length = reader.ReadCodeLength();
                int count = reader.ReadVarint();
                var first = new uint[count];
                var last = new uint[count];
                var cid = new int[count];
                long nextCode = 0L;
                long nextCid = 0L;
                for (int index = 0; index < count; index++) {
                    long start = nextCode + reader.ReadVarintLong();
                    long end = start + reader.ReadVarintLong();
                    long value = nextCid + reader.ReadZigZag();
                    if (end > (1L << (8 * length)) - 1 || value < 0 || value > ushort.MaxValue) {
                        throw new InvalidDataException("Invalid embedded Adobe CMap range.");
                    }
                    first[index] = (uint)start;
                    last[index] = (uint)end;
                    cid[index] = (int)value;
                    nextCode = end + 1;
                    nextCid = constant ? value : value + (end - start) + 1;
                }
                runs[length] = new CodeRuns(first, last, cid, constant);
            }
            return runs;
        }

        private static CidUnicodeMap GetCollectionUnicode(string ordering) =>
            CollectionUnicode.GetOrAdd(ordering, static key => new Lazy<CidUnicodeMap>(() => LoadUnicode(key))).Value;

        private static CidUnicodeMap LoadUnicode(string ordering) {
            var reader = new PayloadReader(ReadEntry(ordering, UnicodeEntry));
            int cidCount = reader.ReadVarint();
            if (cidCount is <= 0 or > ushort.MaxValue + 1) throw new InvalidDataException("Invalid embedded Adobe Unicode map.");
            var scalars = new int[cidCount];
            Dictionary<int, string>? sequences = null;
            int entries = reader.ReadVarint();
            long nextCid = 0L;
            long previousScalar = 0L;
            for (int entry = 0; entry < entries; entry++) {
                long cid = nextCid + reader.ReadVarint();
                int header = reader.ReadVarint();
                if ((header & 1) != 0) {
                    var text = new StringBuilder();
                    for (int index = 0; index < header >> 1; index++) {
                        previousScalar += reader.ReadZigZag();
                        AppendScalar(text, previousScalar);
                    }
                    if (cid >= cidCount) throw new InvalidDataException("Invalid embedded Adobe Unicode map.");
                    (sequences ??= new Dictionary<int, string>())[(int)cid] = text.ToString();
                    scalars[cid] = -1;
                    nextCid = cid + 1;
                } else {
                    int count = (header >> 1) + 1;
                    long scalar = previousScalar + reader.ReadZigZag();
                    if (cid + count > cidCount || scalar < 0 || scalar + count - 1 > 0x10FFFF) {
                        throw new InvalidDataException("Invalid embedded Adobe Unicode map.");
                    }
                    for (int offset = 0; offset < count; offset++) scalars[cid + offset] = (int)scalar + offset;
                    nextCid = cid + count;
                    previousScalar = scalar + count - 1;
                }
            }
            if (!reader.AtEnd) throw new InvalidDataException("Invalid embedded Adobe Unicode map.");
            return new CidUnicodeMap(scalars, sequences);
        }

        private static void AppendScalar(StringBuilder text, long scalar) {
            if (scalar is < 0 or > 0x10FFFF || scalar is >= 0xD800 and <= 0xDFFF) throw new InvalidDataException("Invalid embedded Adobe Unicode value.");
            if (scalar < 0x10000) {
                text.Append((char)scalar);
            } else {
                text.Append(char.ConvertFromUtf32((int)scalar));
            }
        }

        private static byte[] ReadEntry(string ordering, string name) {
            CollectionIndex index = Collections.GetOrAdd(ordering, static key => new Lazy<CollectionIndex>(() => CollectionIndex.Read(key))).Value;
            if (!index.Entries.TryGetValue(name, out CollectionEntry entry)) throw new InvalidDataException("Missing embedded Adobe CMap: " + name);
            using Stream resource = OpenResource(ordering);
            resource.Position = index.DataOffset + entry.Offset;
            var compressed = new byte[entry.CompressedLength];
            ReadExactly(resource, compressed);
            var payload = new byte[entry.Length];
            using (var deflate = new DeflateStream(new MemoryStream(compressed), CompressionMode.Decompress)) {
                ReadExactly(deflate, payload);
                if (deflate.ReadByte() != -1) throw new InvalidDataException("Invalid embedded Adobe CMap length: " + name);
            }
            return payload;
        }

        private static void ReadExactly(Stream stream, byte[] buffer) {
            for (int offset = 0; offset < buffer.Length;) {
                int read = stream.Read(buffer, offset, buffer.Length - offset);
                if (read <= 0) throw new InvalidDataException("Truncated embedded Adobe CMap data.");
                offset += read;
            }
        }

        private static Stream OpenResource(string ordering) => typeof(PdfPredefinedCMap).Assembly.GetManifestResourceStream(
            "OfficeIMO.Pdf.AdobeCMaps.Adobe-" + ordering + ".bin") ?? throw new InvalidDataException("Missing embedded Adobe CMap collection: " + ordering);

        private readonly struct CollectionEntry {
            internal CollectionEntry(int offset, int compressedLength, int length) {
                Offset = offset;
                CompressedLength = compressedLength;
                Length = length;
            }

            internal int Offset { get; }
            internal int CompressedLength { get; }
            internal int Length { get; }
        }

        private sealed class CollectionIndex {
            private static readonly byte[] Magic = { (byte)'O', (byte)'C', (byte)'M', (byte)'P' };

            private CollectionIndex(long dataOffset, Dictionary<string, CollectionEntry> entries) {
                DataOffset = dataOffset;
                Entries = entries;
            }

            internal long DataOffset { get; }
            internal Dictionary<string, CollectionEntry> Entries { get; }

            internal static CollectionIndex Read(string ordering) {
                using Stream resource = OpenResource(ordering);
                var magic = new byte[Magic.Length];
                ReadExactly(resource, magic);
                if (!magic.SequenceEqual(Magic)) throw new InvalidDataException("Invalid embedded Adobe CMap collection: " + ordering);
                int indexLength = 0;
                for (int shift = 0; ; shift += 7) {
                    int value = resource.ReadByte();
                    if (value < 0 || shift > 28) throw new InvalidDataException("Invalid embedded Adobe CMap collection: " + ordering);
                    indexLength |= (value & 0x7F) << shift;
                    if ((value & 0x80) == 0) break;
                }
                var index = new byte[indexLength];
                ReadExactly(resource, index);
                var reader = new PayloadReader(index);
                if (reader.ReadVarint() != 1) throw new InvalidDataException("Unsupported embedded Adobe CMap format: " + ordering);
                reader.ReadText(); // Copyright notice carried with the data.
                int count = reader.ReadVarint();
                var entries = new Dictionary<string, CollectionEntry>(count, StringComparer.Ordinal);
                for (int entry = 0; entry < count; entry++) {
                    string name = reader.ReadText();
                    entries[name] = new CollectionEntry(reader.ReadVarint(), reader.ReadVarint(), reader.ReadVarint());
                }
                if (!reader.AtEnd) throw new InvalidDataException("Invalid embedded Adobe CMap collection: " + ordering);
                return new CollectionIndex(resource.Position, entries);
            }
        }

        private sealed class PayloadReader {
            private readonly byte[] _data;
            private int _position;

            internal PayloadReader(byte[] data) => _data = data;

            internal bool AtEnd => _position == _data.Length;

            internal int ReadByte() {
                if (_position >= _data.Length) throw new InvalidDataException("Truncated embedded Adobe CMap data.");
                return _data[_position++];
            }

            internal int ReadVarint() {
                int value = 0;
                for (int shift = 0; ; shift += 7) {
                    int current = ReadByte();
                    if (shift > 28) throw new InvalidDataException("Invalid embedded Adobe CMap number.");
                    value |= (current & 0x7F) << shift;
                    if ((current & 0x80) == 0) return value < 0 ? throw new InvalidDataException("Invalid embedded Adobe CMap number.") : value;
                }
            }

            internal long ReadVarintLong() {
                long value = 0L;
                for (int shift = 0; ; shift += 7) {
                    int current = ReadByte();
                    if (shift > 56) throw new InvalidDataException("Invalid embedded Adobe CMap number.");
                    value |= (long)(current & 0x7F) << shift;
                    if ((current & 0x80) == 0) return value;
                }
            }

            internal long ReadZigZag() {
                long value = (uint)ReadVarint();
                return (value & 1) == 0 ? value >> 1 : -((value + 1) >> 1);
            }

            internal int ReadCodeLength() {
                int length = ReadByte();
                return length is >= 1 and <= MaximumCodeLength ? length : throw new InvalidDataException("Invalid embedded Adobe CMap code length.");
            }

            internal byte[] ReadBytes(int count) {
                if (_position + count > _data.Length) throw new InvalidDataException("Truncated embedded Adobe CMap data.");
                var bytes = new byte[count];
                Buffer.BlockCopy(_data, _position, bytes, 0, count);
                _position += count;
                return bytes;
            }

            internal string ReadText() => Encoding.ASCII.GetString(ReadBytes(ReadVarint()));
        }

        // One codespace range; each byte position is bounded independently (ISO 32000-2, 9.7.6.2).
        private sealed class CodespaceRange {
            private readonly byte[] _low;
            private readonly byte[] _high;

            internal CodespaceRange(int length, byte[] low, byte[] high) {
                Length = length;
                _low = low;
                _high = high;
            }

            internal int Length { get; }

            internal int MatchingPrefix(byte[] bytes, int offset, int available) {
                int index = 0;
                while (index < available && bytes[offset + index] >= _low[index] && bytes[offset + index] <= _high[index]) index++;
                return index;
            }
        }

        // Sorted, non-overlapping code runs of one code length.
        private sealed class CodeRuns {
            private readonly uint[] _first;
            private readonly uint[] _last;
            private readonly int[] _cid;
            private readonly bool _constant;

            internal CodeRuns(uint[] first, uint[] last, int[] cid, bool constant) {
                _first = first;
                _last = last;
                _cid = cid;
                _constant = constant;
            }

            internal bool TryGet(uint code, out int cid) {
                int low = 0;
                int high = _first.Length - 1;
                while (low <= high) {
                    int middle = low + ((high - low) >> 1);
                    if (code < _first[middle]) {
                        high = middle - 1;
                    } else if (code > _last[middle]) {
                        low = middle + 1;
                    } else {
                        cid = _constant ? _cid[middle] : _cid[middle] + (int)(code - _first[middle]);
                        return true;
                    }
                }
                cid = 0;
                return false;
            }
        }

        // CID-to-Unicode values of one character collection. -1 marks a CID with a multi-scalar value.
        private sealed class CidUnicodeMap {
            private readonly int[] _scalars;
            private readonly Dictionary<int, string>? _sequences;

            internal CidUnicodeMap(int[] scalars, Dictionary<int, string>? sequences) {
                _scalars = scalars;
                _sequences = sequences;
            }

            internal bool TryAppend(int cid, StringBuilder text) {
                if (cid <= 0 || cid >= _scalars.Length) return false;
                int scalar = _scalars[cid];
                if (scalar == 0) return false;
                if (scalar < 0) {
                    text.Append(_sequences![cid]);
                } else if (scalar < 0x10000) {
                    text.Append((char)scalar);
                } else {
                    text.Append(char.ConvertFromUtf32(scalar));
                }
                return true;
            }

            internal IEnumerable<(int Cid, int Scalar)> EnumerateScalars() {
                for (int cid = 1; cid < _scalars.Length; cid++) {
                    if (_scalars[cid] != 0) yield return (cid, _scalars[cid]);
                }
            }
        }
    }
}
