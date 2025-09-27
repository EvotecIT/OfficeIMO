using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static class PdfSyntax {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
#if NET8_0_OR_GREATER
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+(\d+)\s+obj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
#else
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+(\d+)\s+obj", RegexOptions.Compiled, RegexTimeout);
#endif

    internal static (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) ParseObjects(byte[] pdf) {
        string text = PdfEncoding.Latin1GetString(pdf);
        var map = new Dictionary<int, PdfIndirectObject>();
        var matches = ObjRegex.Matches(text);
        for (int i = 0; i < matches.Count; i++) {
            int id = int.Parse(matches[i].Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            int gen = int.Parse(matches[i].Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
            int start = matches[i].Index;
            int end = (i + 1 < matches.Count) ? matches[i + 1].Index : text.Length;
            string body = text.Substring(start, end - start);
            if (TryExtractDictionary(body, out var dictText, out _, out int dictEnd)) {
                var dict = ParseDictionary(dictText);
                int streamIdx = IndexOfKeyword(body, "stream", dictEnd);
                if (streamIdx >= 0) {
                    int dataStart = streamIdx + "stream".Length;
                    // Skip optional whitespace before the newline separating the stream keyword and data
                    while (dataStart < body.Length && (body[dataStart] == ' ' || body[dataStart] == '\t')) dataStart++;
                    if (dataStart < body.Length && body[dataStart] == '\r') {
                        dataStart++;
                        if (dataStart < body.Length && body[dataStart] == '\n') dataStart++;
                    } else if (dataStart < body.Length && body[dataStart] == '\n') {
                        dataStart++;
                    }

                    int endStreamIdx = IndexOfKeyword(body, "endstream", dataStart);
                    if (endStreamIdx >= 0) {
                        int dataEnd = endStreamIdx;
                        // Trim trailing newline characters before endstream to mirror previous regex behaviour
                        while (dataEnd > dataStart && (body[dataEnd - 1] == '\r' || body[dataEnd - 1] == '\n')) dataEnd--;

                        string streamText = dataEnd > dataStart ? body.Substring(dataStart, dataEnd - dataStart) : string.Empty;
                        var data = PdfEncoding.Latin1GetBytes(streamText);
                        // Handle FlateDecode (best-effort, zero-dep)
                        if (HasFlateDecode(dict)) {
                            try { data = Filters.FlateDecoder.Decode(data); } catch (Exception ex) {
                                // Provide failure feedback while keeping original bytes
                                System.Diagnostics.Trace.WriteLine($"OfficeIMO.Pdf: FlateDecode failed for object {id} {gen} R: {ex.Message}");
                                map[id] = new PdfIndirectObject(id, gen, new PdfStream(dict, data, decodingFailed: true, error: ex.Message));
                                continue;
                            }
                        }
                        map[id] = new PdfIndirectObject(id, gen, new PdfStream(dict, data));
                        continue;
                    }
                }

                map[id] = new PdfIndirectObject(id, gen, dict);
            }
        }
        int trailerIdx = text.LastIndexOf("trailer", StringComparison.OrdinalIgnoreCase);
        string trailerRaw = trailerIdx >= 0 ? text.Substring(trailerIdx) : string.Empty;
        return (map, trailerRaw);
    }

    private static bool HasFlateDecode(PdfDictionary dict) {
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
        string tok = tokens[i];
        if (tok == "[") {
            var arr = new PdfArray(); int j = i + 1;
            while (j < tokens.Count && tokens[j] != "]") {
                var (inner, used) = ParseObject(tokens, j);
                arr.Items.Add(inner);
                j += used + 1;
            }
            return (arr, j - i);
        }
        if (tok == "<<") {
            var (dict, consumed) = ParseDictionaryTokens(tokens, i);
            return (dict, consumed);
        }
        if (tok.Length > 0 && tok[0] == '/') return (new PdfName(tok.Substring(1)), 0);
        if (tok.Length > 0 && tok[0] == '(') return (new PdfStringObj(Unescape(tok.Substring(1, tok.Length - 2))), 0);
        if (char.IsDigit(tok[0]) || tok[0] == '-' || tok[0] == '+') {
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
        var tokens = new List<string>();
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
            i = j;
        }
        return tokens;
    }

    private static string Unescape(string s) => PdfTextExtractor.UnescapePdfLiteral(s);

    private static (PdfDictionary Dict, int Consumed) ParseDictionaryTokens(List<string> tokens, int start) {
        var dict = new PdfDictionary();
        int depth = 1;
        int i = start + 1;
        while (i < tokens.Count) {
            string token = tokens[i];
            if (token == ">>") {
                depth--;
                if (depth == 0) return (dict, i - start);
                i++;
                continue;
            }
            if (token == "<<") {
                depth++;
                i++;
                continue;
            }
            if (token.Length > 0 && token[0] == '/') {
                string key = token.Substring(1);
                if (i + 1 < tokens.Count) {
                    if (tokens[i + 1] == "<<") depth++;
                    var (value, used) = ParseObject(tokens, i + 1);
                    dict.Items[key] = value;
                    i += used + 1;
                    continue;
                }
            }
            i++;
        }
        return (dict, Math.Max(0, tokens.Count - start - 1));
    }

    private static bool TryExtractDictionary(string text, out string dictText, out int dictStart, out int dictEnd) {
        dictText = string.Empty;
        dictStart = -1;
        dictEnd = -1;

        int candidate = text.IndexOf("<<", StringComparison.Ordinal);
        if (candidate < 0) return false;

        int depth = 0;
        bool inString = false;
        int stringDepth = 0;
        bool inHexString = false;
        bool inComment = false;

        for (int i = candidate; i < text.Length; i++) {
            char c = text[i];
            char next = i + 1 < text.Length ? text[i + 1] : '\0';

            if (inComment) {
                if (c == '\r' || c == '\n') inComment = false;
                continue;
            }

            if (inString) {
                if (c == '\\') { i++; continue; }
                if (c == '(') { stringDepth++; continue; }
                if (c == ')') {
                    stringDepth--;
                    if (stringDepth <= 0) { inString = false; stringDepth = 0; }
                }
                continue;
            }

            if (inHexString) {
                if (c == '>') inHexString = false;
                continue;
            }

            if (c == '%') { inComment = true; continue; }

            if (c == '(') { inString = true; stringDepth = 1; continue; }

            if (c == '<') {
                if (next == '<') {
                    if (depth == 0) dictStart = i;
                    depth++;
                    i++; // Skip next '<'
                    continue;
                }
                inHexString = true;
                continue;
            }

            if (c == '>' && next == '>') {
                depth--;
                if (depth == 0) {
                    int contentStart = dictStart + 2;
                    int contentLength = i - contentStart;
                    if (contentLength < 0) contentLength = 0;
                    dictText = contentLength > 0 ? text.Substring(contentStart, contentLength) : string.Empty;
                    dictEnd = i + 2;
                    return true;
                }
                i++;
                continue;
            }
        }

        return false;
    }

    private static int IndexOfKeyword(string text, string keyword, int startIndex) {
        int idx = startIndex < 0 ? -1 : startIndex;
        while (idx >= 0 && idx < text.Length) {
            idx = text.IndexOf(keyword, idx, StringComparison.Ordinal);
            if (idx < 0) break;
            bool beforeOk = idx == 0 || char.IsWhiteSpace(text[idx - 1]) || text[idx - 1] == '>';
            int afterIndex = idx + keyword.Length;
            bool afterOk = afterIndex >= text.Length || char.IsWhiteSpace(text[afterIndex]) || text[afterIndex] == '\r' || text[afterIndex] == '\n';
            if (beforeOk && afterOk) return idx;
            idx = afterIndex;
        }
        return -1;
    }
}
