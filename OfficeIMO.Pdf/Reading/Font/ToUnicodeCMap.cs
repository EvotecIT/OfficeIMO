using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal sealed class ToUnicodeCMap {
    private readonly Dictionary<string, string> _map = new(StringComparer.OrdinalIgnoreCase);
    private int _maxKeyBytes = 1;

    public static bool TryParse(byte[] data, out ToUnicodeCMap? cmap) {
        try {
            string s = PdfEncoding.Latin1GetString(data);
            var inst = new ToUnicodeCMap();
            inst.Parse(s);
            cmap = inst; return true;
        } catch { cmap = null; return false; }
    }

    private void Parse(string s) {
        // Handle beginbfchar / endbfchar lines like: <AB> <00E9>
        var bfchar = new Regex(@"<(?<src>[0-9A-Fa-f]+)>\s+<(?<dst>[0-9A-Fa-f]+)>");
        foreach (Match section in Regex.Matches(s, @"beginbfchar([\s\S]*?)endbfchar", RegexOptions.IgnoreCase)) {
            foreach (Match m in bfchar.Matches(section.Groups[1].Value)) {
                AddMap(m.Groups["src"].Value, m.Groups["dst"].Value);
            }
        }
        // Handle beginbfrange / endbfrange, sequential mapping
        var bfrangeLine = new Regex(@"<(?<from>[0-9A-Fa-f]+)>\s+<(?<to>[0-9A-Fa-f]+)>\s+<(?<dst>[0-9A-Fa-f]+)>");
        foreach (Match section in Regex.Matches(s, @"beginbfrange([\s\S]*?)endbfrange", RegexOptions.IgnoreCase)) {
            string body = section.Groups[1].Value;
            foreach (Match m in bfrangeLine.Matches(body)) {
                int from = Convert.ToInt32(m.Groups["from"].Value, 16);
                int to = Convert.ToInt32(m.Groups["to"].Value, 16);
                int dst = Convert.ToInt32(m.Groups["dst"].Value, 16);
                int keyBytes = m.Groups["from"].Value.Length / 2; // bytes per code
                for (int code = from, u = dst; code <= to; code++, u++) {
                    string srcHex = code.ToString("X", System.Globalization.CultureInfo.InvariantCulture).PadLeft(keyBytes * 2, '0');
                    string dstHex = u.ToString("X", System.Globalization.CultureInfo.InvariantCulture);
                    AddMap(srcHex, dstHex);
                }
            }
        }
    }

    private void AddMap(string srcHex, string dstHex) {
        if (srcHex.Length % 2 != 0) srcHex = "0" + srcHex;
        _maxKeyBytes = Math.Max(_maxKeyBytes, srcHex.Length / 2);
        string key = srcHex.ToUpperInvariant();
        // dst may be multi-codepoints; keep as UTF-16 string
        string s = HexToString(dstHex);
        _map[key] = s;
    }

    private static string HexToString(string hex) {
        // Interpret as sequence of 16-bit big-endian code points
        if (hex.Length % 4 != 0) hex = hex.PadLeft(((hex.Length + 3) / 4) * 4, '0');
        var chars = new List<char>(hex.Length / 4);
        for (int i = 0; i < hex.Length; i += 4) {
            ushort code = Convert.ToUInt16(hex.Substring(i, 4), 16);
            chars.Add((char)code);
        }
        return new string(chars.ToArray());
    }

    public string MapBytes(byte[] bytes) {
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < bytes.Length;) {
            // Greedy match up to _maxKeyBytes (1-2 bytes typical)
            int max = Math.Min(_maxKeyBytes, bytes.Length - i);
            string? mapped = null; int used = 0;
            for (int len = max; len >= 1; len--) {
                string key = ByteSliceToHex(bytes, i, len);
                if (_map.TryGetValue(key, out var s)) { mapped = s; used = len; break; }
            }
            if (mapped is null) { sb.Append((char)bytes[i]); i++; } else { sb.Append(mapped); i += used; }
        }
        return sb.ToString();
    }

    private static string ByteSliceToHex(byte[] b, int start, int len) {
        var sb = new System.Text.StringBuilder(len * 2);
        for (int i = 0; i < len; i++) sb.Append(b[start + i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        return sb.ToString();
    }
}
