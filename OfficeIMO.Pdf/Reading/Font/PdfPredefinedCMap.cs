using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf {

    /// <summary>Offline Adobe mappings for the four horizontal UCS-2 CJK encodings.</summary>
    /// <remarks>CID-to-Unicode data comes only from pinned assembly resources. A document-supplied encoding CMap
    /// contributes two-byte code-to-CID ranges alone, parsed within the fixed 65,536-code table.</remarks>
    internal sealed class PdfPredefinedCMap {
        private static readonly char[] Separators = { ' ', '\t' };
        private static readonly Lazy<PdfPredefinedCMap> Japanese = new(() => Load("UniJIS-UCS2-H", "Japan1"));
        private static readonly Lazy<PdfPredefinedCMap> SimplifiedChinese = new(() => Load("UniGB-UCS2-H", "GB1"));
        private static readonly Lazy<PdfPredefinedCMap> TraditionalChinese = new(() => Load("UniCNS-UCS2-H", "CNS1"));
        private static readonly Lazy<PdfPredefinedCMap> Korean = new(() => Load("UniKS-UCS2-H", "Korea1"));
        private static readonly Dictionary<string, Lazy<ToUnicodeCMap>> CollectionUnicode = new(StringComparer.Ordinal) {
            ["Japan1"] = new(() => LoadCollectionUnicode("Japan1")),
            ["GB1"] = new(() => LoadCollectionUnicode("GB1")),
            ["CNS1"] = new(() => LoadCollectionUnicode("CNS1")),
            ["Korea1"] = new(() => LoadCollectionUnicode("Korea1"))
        };
        private static readonly Regex CodespaceSectionRegex = new(@"begincodespacerange([\s\S]*?)endcodespacerange", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex CidRangeSectionRegex = new(@"begincidrange([\s\S]*?)endcidrange", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex CidCharSectionRegex = new(@"begincidchar([\s\S]*?)endcidchar", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex HexCodeRegex = new(@"<\s*([0-9A-Fa-f]{1,8})\s*>", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex CidRangeEntryRegex = new(@"<\s*([0-9A-Fa-f]{4})\s*>\s*<\s*([0-9A-Fa-f]{4})\s*>\s*(\d{1,5})", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private static readonly Regex CidCharEntryRegex = new(@"<\s*([0-9A-Fa-f]{4})\s*>\s*(\d{1,5})", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
        private const long MaximumEmbeddedCMapAssignments = 1_000_000L;

        /// <summary>A map that decodes nothing, used when an embedded CMap cannot be read.</summary>
        internal static PdfPredefinedCMap Empty { get; } = new(new ushort[65536], new ToUnicodeCMap());
        private readonly ushort[] _cids;
        private readonly ToUnicodeCMap _unicode;

        private PdfPredefinedCMap(ushort[] cids, ToUnicodeCMap unicode) {
            _cids = cids;
            _unicode = unicode;
        }

        /// <summary>Returns a shared lazy map only for a compatible Adobe character collection.</summary>
        internal static Lazy<PdfPredefinedCMap>? Find(string encoding, string registry, string ordering) {
            if (!string.Equals(registry, "Adobe", StringComparison.Ordinal)) return null;
            return (encoding, ordering) switch {
                ("UniJIS-UCS2-H", "Japan1") => Japanese,
                ("UniGB-UCS2-H", "GB1") => SimplifiedChinese,
                ("UniCNS-UCS2-H", "CNS1") => TraditionalChinese,
                ("UniKS-UCS2-H", "Korea1") => Korean,
                _ => null
            };
        }

        /// <summary>True when the bundled resources map this character collection's CIDs to Unicode.</summary>
        internal static bool HasUnicodeCollection(string registry, string ordering) =>
            string.Equals(registry, "Adobe", StringComparison.Ordinal) && CollectionUnicode.ContainsKey(ordering);

        /// <summary>
        /// Reads the two-byte code-to-CID ranges of an embedded encoding CMap. Returns <see cref="Empty"/> when the
        /// CMap uses other code lengths, a usecmap parent, or malformed ranges, so text keeps the unmapped diagnostic.
        /// </summary>
        internal static PdfPredefinedCMap FromEmbeddedCMap(byte[] data, string ordering) {
            if (!CollectionUnicode.TryGetValue(ordering, out Lazy<ToUnicodeCMap>? unicode)) return Empty;
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
            return new PdfPredefinedCMap(cids, unicode.Value);
        }

        /// <summary>Maps a complete two-byte code to its CID; prefixes and undefined codes have no mapping.</summary>
        internal bool TryGetCid(byte[] bytes, int offset, out ushort cid) {
            cid = 0;
            if (offset < 0 || offset + 1 >= bytes.Length) return false;
            int code = (bytes[offset] << 8) | bytes[offset + 1];
            cid = _cids[code];
            return cid != 0;
        }

        /// <summary>Resolves a valid painted code, including the encoding's default CID 0.</summary>
        internal bool TryGetPaintedCid(byte[] bytes, int offset, out ushort cid) {
            cid = 0;
            if (offset < 0 || offset + 1 >= bytes.Length) return false;
            int code = (bytes[offset] << 8) | bytes[offset + 1];
            if (code is >= 0xD800 and <= 0xDFFF) return false;
            cid = _cids[code];
            return true;
        }

        /// <summary>Decodes every painted code through both authoritative maps, refusing incomplete coverage.</summary>
        internal bool TryDecode(byte[] bytes, int maximumCharacters, out string decoded) {
            decoded = string.Empty;
            if (bytes.Length % 2 != 0) return false;
            byte[] cidBytes = new byte[bytes.Length];
            for (int index = 0; index < bytes.Length; index += 2) {
                if (!TryGetCid(bytes, index, out ushort cid)) return false;
                cidBytes[index] = (byte)(cid >> 8);
                cidBytes[index + 1] = (byte)cid;
            }
            return _unicode.TryMapBytes(cidBytes, maximumCharacters, out decoded);
        }

        private static PdfPredefinedCMap Load(string encoding, string ordering) {
            ushort[] cids = new ushort[65536];
            using (Stream data = OpenResource(encoding))
            using (StreamReader reader = new StreamReader(data, Encoding.ASCII)) {
                bool inRange = false;
                bool notdefRange = false;
                string? line;
                while ((line = reader.ReadLine()) != null) {
                    if (line.EndsWith(" begincidrange", StringComparison.Ordinal)) { inRange = true; notdefRange = false; continue; }
                    if (line.EndsWith(" beginnotdefrange", StringComparison.Ordinal)) { inRange = true; notdefRange = true; continue; }
                    if (line is "endcidrange" or "endnotdefrange") { inRange = false; continue; }
                    if (!inRange) continue;
                    string[] values = line.Split(Separators, StringSplitOptions.RemoveEmptyEntries);
                    if (values.Length != 3) throw new InvalidDataException("Invalid embedded Adobe CID range.");
                    int first = ParseHexCode(values[0]);
                    int last = ParseHexCode(values[1]);
                    int cid = int.Parse(values[2], CultureInfo.InvariantCulture);
                    if (first > last || cid <= 0 || (long)cid + (notdefRange ? 0 : last - first) > ushort.MaxValue) {
                        throw new InvalidDataException("Invalid embedded Adobe CID range bounds.");
                    }
                    for (int code = first; code <= last; code++) {
                        if (!notdefRange || cids[code] == 0) cids[code] = (ushort)(cid + (notdefRange ? 0 : code - first));
                    }
                }
            }
            return new PdfPredefinedCMap(cids, CollectionUnicode[ordering].Value);
        }

        private static ToUnicodeCMap LoadCollectionUnicode(string ordering) {
            using Stream unicodeData = OpenResource("Adobe-" + ordering + "-UCS2");
            using MemoryStream output = new MemoryStream();
            unicodeData.CopyTo(output);
            if (!ToUnicodeCMap.TryParse(output.ToArray(), out ToUnicodeCMap? unicode) || unicode == null || unicode.MappingCount == 0) {
                throw new InvalidDataException("Invalid embedded Adobe Unicode mapping.");
            }
            return unicode;
        }

        private static int ParseHexCode(string value) {
            if (value.Length != 6 || value[0] != '<' || value[5] != '>') {
                throw new InvalidDataException("Invalid embedded Adobe UCS-2 code.");
            }
    #if NET8_0_OR_GREATER
            return ushort.Parse(value.AsSpan(1, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
    #else
            return ushort.Parse(value.Substring(1, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
    #endif
        }

        private static Stream OpenResource(string name) => typeof(PdfPredefinedCMap).Assembly.GetManifestResourceStream(
            "OfficeIMO.Pdf.AdobeCMaps." + name) ?? throw new InvalidDataException("Missing embedded Adobe CMap: " + name);
    }
}
