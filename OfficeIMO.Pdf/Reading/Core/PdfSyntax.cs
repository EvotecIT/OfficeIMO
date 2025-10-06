using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static class PdfSyntax {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
#if NET8_0_OR_GREATER
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+(\d+)\s+obj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"<<(.*?)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
#else
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+(\d+)\s+obj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"<<(.*?)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline, RegexTimeout);
#endif

    internal static (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) ParseObjects(byte[] pdf) {
        string text = PdfEncoding.Latin1GetString(pdf);
        var map = new Dictionary<int, PdfIndirectObject>();
        var matches = ObjRegex.Matches(text);
        for (int i = 0; i < matches.Count; i++) {
            int id = int.Parse(matches[i].Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            int gen = int.Parse(matches[i].Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
            int start = matches[i].Index;
            int end = FindObjectEnd(text, start);
            if (end < 0) end = (i + 1 < matches.Count) ? matches[i + 1].Index : text.Length;

            // Extract dictionary (balanced << >>) within object bounds
            int dictStart = text.IndexOf("<<", start, end - start);
            if (dictStart >= 0) {
                int dictEnd = FindDictEnd(text, dictStart, end);
                if (dictEnd > dictStart) {
                    string dictText = SafeSlice(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000); // cap to 1 MB
                    PdfDictionary dict;
                    try { dict = ParseDictionary(dictText); }
                    catch (OutOfMemoryException) { dict = new PdfDictionary(); }

                    // Check for stream section; prefer dictionary /Length when available
                    int streamKw = IndexOfKeyword(text, "stream", dictEnd, end);
                    if (streamKw >= 0) {
                        int dataStart = SkipEOL(text, streamKw + 6, end);
                        // Try /Length first (inline number only)
                        int byteStart = dataStart;
                        int byteLen = -1;
                        var lenNum = dict.Get<PdfNumber>("Length");
                        if (lenNum is not null) {
                            int L = (int)Math.Max(0, Math.Min(int.MaxValue, lenNum.Value));
                            if (byteStart >= 0 && byteStart + L <= pdf.Length) byteLen = L;
                        }
                        if (byteLen < 0) {
                            int endStream = IndexOfKeyword(text, "endstream", dataStart, end);
                            if (endStream > dataStart) byteLen = endStream - dataStart;
                        }
                        if (byteLen >= 0) {
                            bool isImage = (dict.Get<PdfName>("Subtype")?.Name == "Image") || (dict.Get<PdfName>("Type")?.Name == "XObject" && dict.Get<PdfName>("Subtype")?.Name == "Image");
                            if (!isImage) {
                                if (byteStart >= 0 && byteLen >= 0 && byteStart + byteLen <= pdf.Length) {
                                    var data = new byte[byteLen];
                                    Buffer.BlockCopy(pdf, byteStart, data, 0, byteLen);
                                    map[id] = new PdfIndirectObject(id, gen, new PdfStream(dict, data));
                                    continue;
                                }
                            } else {
                                map[id] = new PdfIndirectObject(id, gen, new PdfStream(dict, Array.Empty<byte>()));
                                continue;
                            }
                        }
                    }
                    // No stream; store dictionary-only object
                    map[id] = new PdfIndirectObject(id, gen, dict);
                }
            }
        }
        // Expand object streams (/Type /ObjStm) to populate embedded objects (pages and resources often live there)
        ExpandObjectStreams(map, pdf);
        // Debug: count key object types after expansion
        int pageDicts = 0, catalogs = 0, pagesNodes = 0;
        foreach (var kv in map) {
            if (kv.Value.Value is PdfDictionary d) {
                var t = d.Get<PdfName>("Type")?.Name;
                if (t == "Page") pageDicts++;
                else if (t == "Catalog") catalogs++;
                else if (t == "Pages") pagesNodes++;
            }
        }
        System.Console.WriteLine($"Parsed objects: {map.Count}; Catalog: {catalogs}, Pages nodes: {pagesNodes}, Page dicts: {pageDicts}");
        int trailerIdx = text.LastIndexOf("trailer", StringComparison.OrdinalIgnoreCase);
        string trailerRaw = trailerIdx >= 0 ? text.Substring(trailerIdx) : string.Empty;
        return (map, trailerRaw);
    }

    private static void ExpandObjectStreams(Dictionary<int, PdfIndirectObject> map, byte[] pdf) {
        // Snapshot keys to avoid modifying during enumeration
        var keys = new List<int>(map.Keys);
        int objStmCount = 0, expanded = 0;
        foreach (var id in keys) {
            if (!map.TryGetValue(id, out var ind)) continue;
            if (ind.Value is not PdfStream s) continue;
            var type = s.Dictionary.Get<PdfName>("Type")?.Name;
            if (!string.Equals(type, "ObjStm", StringComparison.Ordinal)) continue;
            objStmCount++;

            // Decode object stream bytes (flate only for now)
            var data = HasFlateDecode(s.Dictionary) ? Filters.FlateDecoder.Decode(s.Data) : s.Data;
            int n = (int)(s.Dictionary.Get<PdfNumber>("N")?.Value ?? 0);
            int first = (int)(s.Dictionary.Get<PdfNumber>("First")?.Value ?? 0);
            if (n <= 0 || first <= 0 || first > data.Length) continue;
            // Header: pairs of objectNumber and offset (ASCII)
            var headerBytes = new byte[first];
            Buffer.BlockCopy(data, 0, headerBytes, 0, first);
            string header = PdfEncoding.Latin1GetString(headerBytes);
            var pairs = ParsePairs(header, n);
            if (pairs.Count != n) continue;
            for (int i = 0; i < n; i++) {
                int objNum = pairs[i].Obj;
                int off = pairs[i].Off;
                int start = first + off;
                int end = (i + 1 < n) ? first + pairs[i + 1].Off : data.Length;
                if (start < 0 || end > data.Length || end <= start) continue;
                int len = end - start;
                var sliceBytes = new byte[len];
                Buffer.BlockCopy(data, start, sliceBytes, 0, len);
                var slice = PdfEncoding.Latin1GetString(sliceBytes);
                var parsed = ParseTopLevelObject(slice);
                if (parsed is not null) { map[objNum] = new PdfIndirectObject(objNum, 0, parsed); expanded++; }
            }
        }
        System.Console.WriteLine($"ObjStm found: {objStmCount}, expanded objects: {expanded}");
    }

    private static List<(int Obj, int Off)> ParsePairs(string header, int n) {
        var list = new List<(int, int)>(n);
        int i = 0; int count = 0;
        while (i < header.Length && count < n) {
            SkipWs();
            if (!ReadInt(out int obj)) break;
            SkipWs();
            if (!ReadInt(out int off)) break;
            list.Add((obj, off)); count++;
        }
        return list;

        void SkipWs() { while (i < header.Length && char.IsWhiteSpace(header[i])) i++; }
        bool ReadInt(out int val) {
            int sign = 1; if (i < header.Length && header[i] == '-') { sign = -1; i++; }
            int start = i; long v = 0; bool any = false;
            while (i < header.Length && char.IsDigit(header[i])) { v = v * 10 + (header[i] - '0'); i++; any = true; if (i - start > 10) break; }
            val = any ? (int)(v * sign) : 0; return any;
        }
    }

    private static PdfObject? ParseTopLevelObject(string body) {
        if (string.IsNullOrWhiteSpace(body)) return null;
        var s = body.TrimStart();
        if (s.StartsWith("<<")) {
            // Find matching >> and parse inside
            int dictStart = body.IndexOf("<<", StringComparison.Ordinal);
            if (dictStart >= 0) {
                int dictEnd = FindDictEnd(body, dictStart, body.Length);
                if (dictEnd > dictStart) {
                    string dictText = SafeSlice(body, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000);
                    try { return ParseDictionary(dictText); } catch { return new PdfDictionary(); }
                }
            }
            return new PdfDictionary();
        }
        if (s.StartsWith("[")) {
            var toks = Tokenize(s);
            var (obj, _) = ParseObject(toks, 0);
            return obj;
        }
        if (s.StartsWith("(")) {
            // literal string
            int end = s.LastIndexOf(')');
            string inner = end > 1 ? s.Substring(1, end - 1) : s.Substring(1);
            return new PdfStringObj(Unescape(inner));
        }
        // number or name fallbacks
        var tokens = Tokenize(s);
        if (tokens.Count > 0) {
            var (obj0, _) = ParseObject(tokens, 0);
            return obj0;
        }
        return null;
    }

    internal static bool HasFlateDecode(PdfDictionary dict) {
        if (!dict.Items.TryGetValue("Filter", out var f)) return false;
        if (f is PdfName n) return string.Equals(n.Name, "FlateDecode", System.StringComparison.Ordinal);
        if (f is PdfArray arr) {
            foreach (var item in arr.Items) if (item is PdfName nn && string.Equals(nn.Name, "FlateDecode", System.StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private static PdfDictionary ParseDictionary(string dict) {
        var d = new PdfDictionary();
        var tokens = Tokenize(dict);
        for (int i = 0; i < tokens.Count; i++) {
            if (tokens[i].Length > 0 && tokens[i][0] == '/') {
                string key = tokens[i].Substring(1);
                if (i + 1 < tokens.Count) {
                    var (obj, consumed) = ParseObject(tokens, i + 1);
                    d.Items[key] = obj;
                    i += consumed;
                }
            }
        }
        return d;
    }

    private static (PdfObject Obj, int Consumed) ParseObject(List<string> tokens, int i) {
        if (i < 0 || i >= tokens.Count) return (new PdfName(""), 0);
        string tok = tokens[i] ?? string.Empty;
        if (tok == "[") {
            var arr = new PdfArray(); int j = i + 1;
            while (j < tokens.Count && tokens[j] != "]") {
                var (inner, used) = ParseObject(tokens, j);
                arr.Items.Add(inner);
                j += used + 1;
            }
            return (arr, j - i);
        }
        if (tok.Length > 0 && tok[0] == '/') return (new PdfName(tok.Substring(1)), 0);
        if (tok.Length > 0 && tok[0] == '(') return (new PdfStringObj(Unescape(tok.Substring(1, tok.Length - 2))), 0);
        if (tok.Length > 0 && (char.IsDigit(tok[0]) || tok[0] == '-' || tok[0] == '+')) {
            // reference (obj gen R) or number
            if (i + 2 < tokens.Count && tokens[i + 2] == "R" && int.TryParse(tokens[i], out int obj) && int.TryParse(tokens[i + 1], out int gen)) {
                return (new PdfReference(obj, gen), 2);
            }
            if (double.TryParse(tok, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val)) {
                return (new PdfNumber(val), 0);
            }
        }
        return (new PdfName(tok), 0);
    }

    private static List<string> Tokenize(string s) {
        // Guardrails for pathological inputs; dictionaries should be small.
        if (s.Length > 1_000_000) s = s.Substring(0, 1_000_000);
        var tokens = new List<string>(Math.Min(16384, s.Length / 2 + 8));
        int i = 0;
        while (i < s.Length) {
            char c = s[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }
            if (c == '<' && i + 1 < s.Length && s[i + 1] == '<') { tokens.Add("<<"); i += 2; continue; }
            if (c == '>' && i + 1 < s.Length && s[i + 1] == '>') { tokens.Add(">>"); i += 2; continue; }
            if (c == '[' || c == ']') { tokens.Add(c.ToString()); i++; continue; }
            if (c == '(') {
                int start = i; i++;
                int depth = 1; bool esc = false;
                var sb = new StringBuilder();
                while (i < s.Length && depth > 0) {
                    char ch = s[i++];
                    if (esc) { sb.Append(ch); esc = false; } else if (ch == '\\') esc = true;
                    else if (ch == '(') { depth++; sb.Append(ch); } else if (ch == ')') { depth--; if (depth > 0) sb.Append(ch); } else sb.Append(ch);
                }
                tokens.Add("(" + sb.ToString() + ")");
                continue;
            }
            // name, number, keyword
            int j = i;
            while (j < s.Length && !char.IsWhiteSpace(s[j]) && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
            string tok = s.Substring(i, j - i);
            if (tok.Length == 0 && s[i] == '/') { // name starting here
                j = i + 1; while (j < s.Length && !char.IsWhiteSpace(s[j]) && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
                tok = s.Substring(i, j - i);
            }
            tokens.Add(tok);
            if (tokens.Count > 100_000) break; // hard stop
            i = j;
        }
        return tokens;
    }

    private static string Unescape(string s) => PdfTextExtractor.UnescapePdfLiteral(s);

    private static int FindObjectEnd(string text, int start) {
        int idx = text.IndexOf("endobj", start, StringComparison.Ordinal);
        return idx >= 0 ? idx + 6 : -1;
    }

    private static int FindDictEnd(string text, int dictStart, int limit) {
        int depth = 0;
        for (int i = dictStart; i + 1 < limit; i++) {
            char c = text[i]; char n = text[i + 1];
            if (c == '<' && n == '<') { depth++; i++; continue; }
            if (c == '>' && n == '>') { depth--; i++; if (depth == 0) return i + 1; continue; }
        }
        return -1;
    }

    private static int IndexOfKeyword(string text, string keyword, int start, int limit) {
        if (start < 0) start = 0; if (limit > text.Length) limit = text.Length;
        int idx = text.IndexOf(keyword, start, StringComparison.Ordinal);
        return (idx >= 0 && idx < limit) ? idx : -1;
    }

    private static int SkipEOL(string text, int idx, int limit) {
        if (idx < limit) {
            if (text[idx] == '\r') idx++;
            if (idx < limit && text[idx] == '\n') idx++;
        }
        return idx;
    }

    private static string SafeSlice(string s, int start, int length, int maxLen) {
        int len = Math.Min(length, Math.Max(0, maxLen));
        if (start < 0) start = 0;
        if (start + len > s.Length) len = s.Length - start;
        if (len <= 0) return string.Empty;
        return s.Substring(start, len);
    }
}
