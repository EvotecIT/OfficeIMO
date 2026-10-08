using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Offline Adobe mappings for the four horizontal UCS-2 CJK encodings.</summary>
/// <remarks>Only pinned assembly resources are parsed. Document-supplied CMaps do not enter this reader.</remarks>
internal sealed class PdfPredefinedCMap {
    private static readonly char[] Separators = { ' ', '\t' };
    private static readonly Lazy<PdfPredefinedCMap> Japanese = new(() => Load("UniJIS-UCS2-H", "Japan1"));
    private static readonly Lazy<PdfPredefinedCMap> SimplifiedChinese = new(() => Load("UniGB-UCS2-H", "GB1"));
    private static readonly Lazy<PdfPredefinedCMap> TraditionalChinese = new(() => Load("UniCNS-UCS2-H", "CNS1"));
    private static readonly Lazy<PdfPredefinedCMap> Korean = new(() => Load("UniKS-UCS2-H", "Korea1"));
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

    /// <summary>Maps a complete two-byte code to its CID; prefixes and undefined codes have no mapping.</summary>
    internal bool TryGetCid(byte[] bytes, int offset, out ushort cid) {
        cid = 0;
        if (offset < 0 || offset + 1 >= bytes.Length) return false;
        int code = (bytes[offset] << 8) | bytes[offset + 1];
        cid = _cids[code];
        return cid != 0;
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

    /// <summary>Checks painted-code coverage independently of an explicit ToUnicode override.</summary>
    internal bool CanMapCodes(byte[] bytes) {
        if (bytes.Length % 2 != 0) return false;
        for (int index = 0; index < bytes.Length; index += 2) {
            if (!TryGetCid(bytes, index, out _)) return false;
        }
        return true;
    }

    private static PdfPredefinedCMap Load(string encoding, string ordering) {
        ushort[] cids = new ushort[65536];
        using (Stream data = OpenResource(encoding))
        using (StreamReader reader = new StreamReader(data, Encoding.ASCII)) {
            bool inRange = false;
            string? line;
            while ((line = reader.ReadLine()) != null) {
                if (line.EndsWith(" begincidrange", StringComparison.Ordinal)) { inRange = true; continue; }
                if (line == "endcidrange") { inRange = false; continue; }
                if (!inRange) continue;
                string[] values = line.Split(Separators, StringSplitOptions.RemoveEmptyEntries);
                if (values.Length != 3) throw new InvalidDataException("Invalid embedded Adobe CID range.");
                int first = ParseHexCode(values[0]);
                int last = ParseHexCode(values[1]);
                int cid = int.Parse(values[2], CultureInfo.InvariantCulture);
                if (first > last || cid <= 0 || (long)cid + last - first > ushort.MaxValue) {
                    throw new InvalidDataException("Invalid embedded Adobe CID range bounds.");
                }
                for (int code = first; code <= last; code++) cids[code] = (ushort)(cid + code - first);
            }
        }
        using Stream unicodeData = OpenResource("Adobe-" + ordering + "-UCS2");
        using MemoryStream output = new MemoryStream();
        unicodeData.CopyTo(output);
        if (!ToUnicodeCMap.TryParse(output.ToArray(), out ToUnicodeCMap? unicode) || unicode == null || unicode.MappingCount == 0) {
            throw new InvalidDataException("Invalid embedded Adobe Unicode mapping.");
        }
        return new PdfPredefinedCMap(cids, unicode);
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
