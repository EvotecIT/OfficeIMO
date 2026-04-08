using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Minimal, zero-dependency text extractor for simple PDFs produced by OfficeIMO.Pdf
/// and common external PDFs with basic text operators and FlateDecode content streams.
/// Not a general-purpose PDF parser; designed as a pragmatic starting point.
/// </summary>
public static class PdfTextExtractor {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
    private static readonly char[] SpaceSplitChars = new[] { ' ' };
#if NET8_0_OR_GREATER
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+0\s+obj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex InfoRefRegex = new Regex(@"/Info\s+(\d+)\s+0\s+R", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex PageObjRegex = new Regex(@"<<(?:.*?)/Type\s*/Page\b(?:.*?)/Contents\s+(?:(?<single>\d+)\s+0\s+R|\[(?<array>[^\]]*)\])(?:.*?)/?>>", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex RefRegex = new Regex(@"(\d+)\s+0\s+R", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex TjRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*Tj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex HexTjRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*Tj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex QuoteLiteralRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*'", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex QuoteHexRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*'", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex DoubleQuoteLiteralRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+\((?<txt>(?:\\.|[^\\\)])*)\)\s*""", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex DoubleQuoteHexRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+<(?<txt>[0-9A-Fa-f\s]+)>\s*""", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
#else
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+0\s+obj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex InfoRefRegex = new Regex(@"/Info\s+(\d+)\s+0\s+R", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex PageObjRegex = new Regex(@"<<(?:.|\n|\r)*?/Type\s*/Page\b(?:.|\n|\r)*?/Contents\s+(?:(?<single>\d+)\s+0\s+R|\[(?<array>[^\]]*)\])(?:.|\n|\r)*?>>", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex RefRegex = new Regex(@"(\d+)\s+0\s+R", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex TjRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*Tj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex HexTjRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*Tj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex QuoteLiteralRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*'", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex QuoteHexRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*'", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex DoubleQuoteLiteralRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+\((?<txt>(?:\\.|[^\\\)])*)\)\s*""", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex DoubleQuoteHexRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+<(?<txt>[0-9A-Fa-f\s]+)>\s*""", RegexOptions.Compiled, RegexTimeout);
#endif

    /// <summary>Extracts plain text from all pages, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(string path) {
        var bytes = File.ReadAllBytes(path);
        return ExtractAllText(bytes);
    }

    /// <summary>Extracts plain text from all pages, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(byte[] pdf) {
        var (parsedObjects, _) = PdfSyntax.ParseObjects(pdf);
        var map = BuildObjectMap(pdf, out _);
        var pages = CollectPages(parsedObjects);
        var sb = new StringBuilder();

        if (pages.Count > 0) {
            for (int i = 0; i < pages.Count; i++) {
                string pageText = ExtractTextFromPage(pages[i], parsedObjects, map);
                if (string.IsNullOrWhiteSpace(pageText)) {
                    continue;
                }

                if (sb.Length > 0) {
                    sb.AppendLine();
                }
                sb.Append(pageText);
            }

            if (sb.Length > 0) {
                return sb.ToString();
            }
        }

        var pageContents = FindPageContentIds(pdf);
        for (int i = 0; i < pageContents.Count; i++) {
            var pageText = new StringBuilder();
            foreach (int contentId in pageContents[i]) {
                if (TryGetContentStreamContent(parsedObjects, map, contentId, out string content)) {
                    pageText.Append(ExtractTextFromContentStream(content));
                }
            }

            if (pageText.Length == 0) {
                continue;
            }

            if (sb.Length > 0) {
                sb.AppendLine();
            }
            sb.Append(pageText);
        }
        return sb.ToString();
    }

    private static string ExtractTextFromPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> parsedObjects, Dictionary<int, string> rawObjects) {
        var pageText = new StringBuilder();
        var resources = ResolveDict(GetInheritedValue(page, "Resources", parsedObjects), parsedObjects);
        var activeForms = new HashSet<int>();

        foreach (int contentId in GetContentIds(page, parsedObjects)) {
            if (TryGetContentStreamContent(parsedObjects, rawObjects, contentId, out string content)) {
                pageText.Append(ExtractTextFromContentStream(content, resources, parsedObjects, rawObjects, activeForms));
            }
        }

        return pageText.ToString();
    }

    private static bool TryGetContentStreamContent(Dictionary<int, PdfIndirectObject> parsedObjects, Dictionary<int, string> rawObjects, int contentId, out string content) {
        if (parsedObjects.TryGetValue(contentId, out var parsedObject) &&
            parsedObject.Value is PdfStream stream) {
            byte[] streamBytes = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, parsedObjects);

            content = PdfEncoding.Latin1GetString(streamBytes);
            return true;
        }

        if (rawObjects.TryGetValue(contentId, out var obj)) {
            var match = StreamRegex.Match(obj);
            if (match.Success) {
                content = match.Groups[1].Value;
                return true;
            }
        }

        content = string.Empty;
        return false;
    }

    /// <summary>Gets document metadata (Title/Author/Subject/Keywords) if present; null when absent.</summary>
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(byte[] pdf) {
        var map = BuildObjectMap(pdf, out var trailer);
        var m = InfoRefRegex.Match(trailer);
        if (!m.Success) return (null, null, null, null);
        int infoId = int.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        if (!map.TryGetValue(infoId, out var obj)) return (null, null, null, null);
        string? title = ExtractStringValue(obj, "/Title");
        string? author = ExtractStringValue(obj, "/Author");
        string? subject = ExtractStringValue(obj, "/Subject");
        string? keywords = ExtractStringValue(obj, "/Keywords");
        return (title, author, subject, keywords);
    }

    private static Dictionary<int, string> BuildObjectMap(byte[] pdf, out string trailer) {
        string text = PdfEncoding.Latin1GetString(pdf);
        var dict = new Dictionary<int, string>();
        var matches = ObjRegex.Matches(text);
        for (int i = 0; i < matches.Count; i++) {
            int id = int.Parse(matches[i].Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            int start = matches[i].Index;
            int end = (i + 1 < matches.Count) ? matches[i + 1].Index : text.Length;
            string body = text.Substring(start, end - start);
            // trim header to just 'obj .. endobj'
            int objStart = body.IndexOf("obj", StringComparison.Ordinal);
            int objEnd = body.IndexOf("endobj", StringComparison.Ordinal);
            if (objStart >= 0 && objEnd > objStart) {
                dict[id] = body.Substring(objStart + 3, objEnd - (objStart + 3));
            }
        }
        int trailerIdx = text.LastIndexOf("trailer", StringComparison.OrdinalIgnoreCase);
        trailer = trailerIdx >= 0 ? text.Substring(trailerIdx) : string.Empty;
        return dict;
    }

    private static List<List<int>> FindPageContentIds(byte[] pdf) {
        string text = PdfEncoding.Latin1GetString(pdf);
        var ids = new List<List<int>>();
        foreach (Match m in PageObjRegex.Matches(text)) {
            if (m.Groups["single"].Success && int.TryParse(m.Groups["single"].Value, out int singleId)) {
                ids.Add(new List<int> { singleId });
                continue;
            }

            if (!m.Groups["array"].Success) {
                continue;
            }

            var pageIds = new List<int>();
            foreach (Match refMatch in RefRegex.Matches(m.Groups["array"].Value)) {
                if (int.TryParse(refMatch.Groups[1].Value, out int id)) {
                    pageIds.Add(id);
                }
            }

            if (pageIds.Count > 0) {
                ids.Add(pageIds);
            }
        }
        return ids;
    }

    private static string ExtractTextFromContentStream(
        string content,
        PdfDictionary? resources = null,
        Dictionary<int, PdfIndirectObject>? parsedObjects = null,
        Dictionary<int, string>? rawObjects = null,
        HashSet<int>? activeForms = null) {
        var sb = new StringBuilder();
        bool inText = false;
        bool pendingSpace = false;
        double currentFontSize = 12;
        double currentHorizontalScale = 1.0;
        var args = new List<object>(8);
        int i = 0;
        int n = content.Length;
        while (i < n) {
            SkipWs();
            if (i >= n) break;

            char c = content[i];
            if (c == '%') {
                while (i < n && content[i] != '\n' && content[i] != '\r') i++;
                continue;
            }

            if (c == '/') { args.Add(ReadName()); continue; }
            if (c == '(') { args.Add(ReadLiteralString()); continue; }
            if (c == '<') {
                if (i + 1 < n && content[i + 1] == '<') { i += 2; continue; }
                args.Add(ReadHexString());
                continue;
            }
            if (c == '[') { args.Add(ReadArray()); continue; }
            if (c == ']' || c == '>') { i++; continue; }
            if (IsNumberStart(c)) { args.Add(ReadNumber()); continue; }

            string op = ReadOperator();
            if (op.Length == 0) {
                i++;
                continue;
            }

            switch (op) {
                case "BT":
                    inText = true;
                    pendingSpace = false;
                    args.Clear();
                    break;
                case "ET":
                    inText = false;
                    pendingSpace = false;
                    args.Clear();
                    break;
                case "T*":
                    if (inText) {
                        sb.AppendLine();
                        pendingSpace = false;
                    }
                    args.Clear();
                    break;
                case "Tf":
                    if (args.Count >= 2) {
                        currentFontSize = ToDouble(args[args.Count - 1]);
                    }
                    args.Clear();
                    break;
                case "Tz":
                    if (args.Count >= 1) {
                        currentHorizontalScale = ToDouble(args[args.Count - 1]) / 100.0;
                    }
                    args.Clear();
                    break;
                case "Td":
                    if (inText && args.Count >= 2) {
                        double advanceX = ToDouble(args[args.Count - 2]);
                        double advanceY = ToDouble(args[args.Count - 1]);
                        if (advanceY == 0 && advanceX > 0.1) {
                            pendingSpace = true;
                        }
                    }
                    args.Clear();
                    break;
                case "Tj":
                    if (inText && args.Count >= 1) {
                        AppendTextRun(ToText(args[args.Count - 1]));
                    }
                    args.Clear();
                    break;
                case "TJ":
                    if (inText && args.Count >= 1) {
                        AppendTextArray(args[args.Count - 1]);
                    }
                    args.Clear();
                    break;
                case "'":
                    if (inText && args.Count >= 1) {
                        RequestSpace();
                        AppendTextRun(ToText(args[args.Count - 1]));
                    }
                    args.Clear();
                    break;
                case "\"":
                    if (inText && args.Count >= 3) {
                        RequestSpace();
                        AppendTextRun(ToText(args[args.Count - 1]));
                    }
                    args.Clear();
                    break;
                case "Do":
                    if (resources is not null && parsedObjects is not null && rawObjects is not null && args.Count >= 1) {
                        string formText = ExtractInvokedFormText(ToName(args[args.Count - 1]), resources, parsedObjects, rawObjects, activeForms ?? new HashSet<int>());
                        if (!string.IsNullOrEmpty(formText)) {
                            AppendTextRun(formText);
                        }
                    }
                    args.Clear();
                    break;
                default:
                    args.Clear();
                    break;
            }
        }

        return sb.ToString();

        void AppendTextRun(string value) {
            if (string.IsNullOrEmpty(value)) return;
            if (pendingSpace &&
                sb.Length > 0 &&
                !char.IsWhiteSpace(sb[sb.Length - 1]) &&
                !char.IsWhiteSpace(value[0])) {
                sb.Append(' ');
            }
            sb.Append(value);
            pendingSpace = false;
        }

        void RequestSpace() {
            pendingSpace = true;
        }

        void AppendTextArray(object arrayObject) {
            if (arrayObject is not List<object> list) {
                return;
            }

            foreach (var item in list) {
                if (item is string text) {
                    AppendTextRun(text);
                } else if (item is double adjustment &&
                           (-adjustment / 1000.0 * currentFontSize * currentHorizontalScale) > Math.Max(1.5, currentFontSize * 0.24)) {
                    RequestSpace();
                }
            }
        }

        void SkipWs() {
            while (i < n && char.IsWhiteSpace(content[i])) i++;
        }

        double ReadNumber() {
            int start = i;
            i++;
            while (i < n) {
                char ch = content[i];
                if (!(char.IsDigit(ch) || ch == '.' || ch == 'E' || ch == 'e' || ch == '-' || ch == '+')) break;
                i++;
            }

            var s = content.Substring(start, i - start);
            if (!double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v)) v = 0;
            return v;
        }

        string ReadName() {
            i++;
            int start = i;
            while (i < n) {
                char ch = content[i];
                if (char.IsWhiteSpace(ch) || ch == '%' || ch == '/' || ch == '[' || ch == ']' || ch == '(' || ch == ')' || ch == '<' || ch == '>') break;
                i++;
            }
            return PdfSyntax.DecodeName(content.Substring(start, i - start));
        }

        string ReadLiteralString() {
            i++;
            int depth = 1;
            bool escaped = false;
            var raw = new StringBuilder();
            while (i < n && depth > 0) {
                char ch = content[i++];
                if (escaped) {
                    raw.Append(ch);
                    escaped = false;
                    continue;
                }

                if (ch == '\\') {
                    raw.Append(ch);
                    escaped = true;
                    continue;
                }

                if (ch == '(') {
                    depth++;
                    raw.Append(ch);
                    continue;
                }

                if (ch == ')') {
                    depth--;
                    if (depth > 0) {
                        raw.Append(ch);
                    }
                    continue;
                }

                raw.Append(ch);
            }

            return UnescapePdfLiteral(raw.ToString());
        }

        string ReadHexString() {
            i++;
            var raw = new StringBuilder();
            while (i < n && content[i] != '>') {
                raw.Append(content[i]);
                i++;
            }

            if (i < n && content[i] == '>') {
                i++;
            }

            return DecodeHexPdfString(raw.ToString());
        }

        List<object> ReadArray() {
            var list = new List<object>();
            i++;
            while (i < n) {
                SkipWs();
                if (i >= n) break;
                char ch = content[i];
                if (ch == ']') {
                    i++;
                    break;
                }

                if (ch == '(') {
                    list.Add(ReadLiteralString());
                    continue;
                }

                if (ch == '<') {
                    if (i + 1 < n && content[i + 1] == '<') {
                        i += 2;
                        continue;
                    }
                    list.Add(ReadHexString());
                    continue;
                }

                if (IsNumberStart(ch)) {
                    list.Add(ReadNumber());
                    continue;
                }

                if (ch == '/') {
                    list.Add(ReadName());
                    continue;
                }

                if (ch == '[') {
                    i++;
                    continue;
                }

                ReadOperator();
            }

            return list;
        }

        string ReadOperator() {
            int start = i;
            char ch = content[i++];
            if (ch == '\'' || ch == '"') return ch.ToString();
            while (i < n) {
                char current = content[i];
                if (char.IsWhiteSpace(current) || current == '%' || current == '(' || current == '[' || current == '/' || current == '<' || current == '>') break;
                i++;
            }
            return content.Substring(start, i - start);
        }

        static bool IsNumberStart(char ch) => ch == '+' || ch == '-' || ch == '.' || char.IsDigit(ch);
        static double ToDouble(object o) => o is double d ? d : 0.0;
        static string ToText(object o) => o as string ?? string.Empty;
        static string ToName(object o) => o as string ?? string.Empty;
    }

    private static string ExtractInvokedFormText(
        string formName,
        PdfDictionary resources,
        Dictionary<int, PdfIndirectObject> parsedObjects,
        Dictionary<int, string> rawObjects,
        HashSet<int> activeForms) {
        if (!TryGetFormStream(resources, formName, parsedObjects, out var formStream, out int formObjectNumber)) {
            return string.Empty;
        }

        bool trackRecursion = formObjectNumber > 0;
        if (trackRecursion && !activeForms.Add(formObjectNumber)) {
            return string.Empty;
        }

        try {
            string content = DecodeStreamContent(formStream, parsedObjects);
            var formResources = ResolveDict(formStream.Dictionary.Items.TryGetValue("Resources", out var resObj) ? resObj : null, parsedObjects) ?? resources;
            return ExtractTextFromContentStream(content, formResources, parsedObjects, rawObjects, activeForms);
        } finally {
            if (trackRecursion) {
                activeForms.Remove(formObjectNumber);
            }
        }
    }

    private static bool TryGetFormStream(
        PdfDictionary resources,
        string name,
        Dictionary<int, PdfIndirectObject> parsedObjects,
        out PdfStream formStream,
        out int objectNumber) {
        formStream = null!;
        objectNumber = 0;

        if (!resources.Items.TryGetValue("XObject", out var xObjectObj)) {
            return false;
        }

        var xObjectDict = ResolveDict(xObjectObj, parsedObjects);
        if (xObjectDict is null || !xObjectDict.Items.TryGetValue(name, out var formObj)) {
            return false;
        }

        if (formObj is PdfReference formRef &&
            parsedObjects.TryGetValue(formRef.ObjectNumber, out var indirectForm) &&
            indirectForm.Value is PdfStream referencedStream &&
            string.Equals(referencedStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = referencedStream;
            objectNumber = formRef.ObjectNumber;
            return true;
        }

        if (formObj is PdfStream directStream &&
            string.Equals(directStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = directStream;
            return true;
        }

        return false;
    }

    private static string DecodeStreamContent(PdfStream stream, Dictionary<int, PdfIndirectObject>? parsedObjects = null) {
        byte[] bytes = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, parsedObjects);
        return PdfEncoding.Latin1GetString(bytes);
    }

    internal static string UnescapePdfLiteral(string s) {
        var sb = new StringBuilder();
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c == '\\' && i + 1 < s.Length) {
                char n = s[++i];
                if (n >= '0' && n <= '7') {
                    int value = n - '0';
                    int digits = 1;
                    while (digits < 3 && i + 1 < s.Length && s[i + 1] >= '0' && s[i + 1] <= '7') {
                        value = (value * 8) + (s[++i] - '0');
                        digits++;
                    }
                    sb.Append((char)value);
                    continue;
                }

                if (n == '\r') {
                    if (i + 1 < s.Length && s[i + 1] == '\n') {
                        i++;
                    }
                    continue;
                }

                if (n == '\n') {
                    continue;
                }

                sb.Append(n switch {
                    'n' => '\n',
                    'r' => '\r',
                    't' => '\t',
                    'b' => '\b',
                    'f' => '\f',
                    '(' => '(',
                    ')' => ')',
                    '\\' => '\\',
                    _ => n
                });
            } else sb.Append(c);
        }
        return sb.ToString();
    }

    internal static string DecodeHexPdfString(string s) {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;

        var hex = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++) {
            char ch = s[i];
            if (!char.IsWhiteSpace(ch)) hex.Append(ch);
        }

        if (hex.Length % 2 == 1) hex.Append('0');

        var bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++) {
            int hi = HexNibble(hex[i * 2]);
            int lo = HexNibble(hex[i * 2 + 1]);
            bytes[i] = (byte)((hi << 4) | lo);
        }

        return PdfWinAnsiEncoding.Decode(bytes);

        static int HexNibble(char c) {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
            if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
            throw new FormatException($"Invalid hex character '{c}'.");
        }
    }

    private static string? ExtractStringValue(string obj, string key) {
        int idx = obj.IndexOf(key, StringComparison.Ordinal);
        if (idx < 0) return null;
        int valueStart = idx + key.Length;
        while (valueStart < obj.Length && char.IsWhiteSpace(obj[valueStart])) {
            valueStart++;
        }

        if (valueStart >= obj.Length) {
            return null;
        }

        if (obj[valueStart] == '(') {
            int close = FindCloseParen(obj, valueStart);
            if (close < 0) return null;
            string raw = obj.Substring(valueStart + 1, close - valueStart - 1);
            return UnescapePdfLiteral(raw);
        }

        if (obj[valueStart] == '<' && (valueStart + 1 >= obj.Length || obj[valueStart + 1] != '<')) {
            int close = obj.IndexOf('>', valueStart + 1);
            if (close < 0) return null;
            string raw = obj.Substring(valueStart + 1, close - valueStart - 1);
            return DecodeMetadataHexString(raw);
        }

        return null;
    }

    private static int FindCloseParen(string s, int start) {
        int depth = 0;
        for (int i = start; i < s.Length; i++) {
            char c = s[i];
            if (c == '\\') { i++; continue; }
            if (c == '(') depth++;
            else if (c == ')') { depth--; if (depth == 0) return i; }
        }
        return -1;
    }

    private static bool TryReadTextAdvance(string line, out double advanceX, out double advanceY) {
        advanceX = 0;
        advanceY = 0;

        if (!line.EndsWith(" Td", StringComparison.Ordinal) && line != "Td") {
            return false;
        }

        var parts = line.Split(SpaceSplitChars, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 3 || !string.Equals(parts[parts.Length - 1], "Td", StringComparison.Ordinal)) {
            return false;
        }

        return double.TryParse(parts[parts.Length - 3], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out advanceX) &&
               double.TryParse(parts[parts.Length - 2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out advanceY);
    }

    private static bool TryReadFontSize(string line, out double fontSize) {
        fontSize = 0;
        var parts = line.Split(SpaceSplitChars, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length >= 3 &&
               string.Equals(parts[parts.Length - 1], "Tf", StringComparison.Ordinal) &&
               double.TryParse(parts[parts.Length - 2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out fontSize);
    }

    private static bool TryReadHorizontalScale(string line, out double horizontalScale) {
        horizontalScale = 0;
        var parts = line.Split(SpaceSplitChars, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2 ||
            !string.Equals(parts[parts.Length - 1], "Tz", StringComparison.Ordinal) ||
            !double.TryParse(parts[parts.Length - 2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double percent)) {
            return false;
        }

        horizontalScale = percent / 100.0;
        return true;
    }

    private static bool TryAppendTextArray(string line, double fontSize, double horizontalScale, Action<string> appendText, Action requestSpace) {
        int operatorIndex = line.LastIndexOf("TJ", StringComparison.Ordinal);
        if (operatorIndex < 0) {
            return false;
        }

        int arrayStart = line.IndexOf('[');
        int arrayEnd = line.LastIndexOf(']');
        if (arrayStart < 0 || arrayEnd < arrayStart || arrayEnd > operatorIndex) {
            return false;
        }

        string items = line.Substring(arrayStart + 1, arrayEnd - arrayStart - 1);
        int i = 0;
        while (i < items.Length) {
            while (i < items.Length && char.IsWhiteSpace(items[i])) {
                i++;
            }

            if (i >= items.Length) {
                break;
            }

            char ch = items[i];
            if (ch == '(') {
                appendText(ReadLiteral(items, ref i));
                continue;
            }

            if (ch == '<') {
                appendText(ReadHex(items, ref i));
                continue;
            }

            if (IsNumberStart(ch)) {
                if (TryReadNumber(items, ref i, out double adjustment) &&
                    ToVisualAdvance(adjustment, fontSize, horizontalScale) > Math.Max(1.5, fontSize * 0.24)) {
                    requestSpace();
                }
                continue;
            }

            i++;
        }

        return true;

        static string ReadLiteral(string s, ref int index) {
            index++; // skip (
            int depth = 1;
            bool escaped = false;
            var raw = new StringBuilder();
            while (index < s.Length && depth > 0) {
                char current = s[index++];
                if (escaped) {
                    raw.Append(current);
                    escaped = false;
                    continue;
                }

                if (current == '\\') {
                    raw.Append(current);
                    escaped = true;
                    continue;
                }

                if (current == '(') {
                    depth++;
                    raw.Append(current);
                    continue;
                }

                if (current == ')') {
                    depth--;
                    if (depth > 0) {
                        raw.Append(current);
                    }
                    continue;
                }

                raw.Append(current);
            }

            return UnescapePdfLiteral(raw.ToString());
        }

        static string ReadHex(string s, ref int index) {
            index++; // skip <
            var raw = new StringBuilder();
            while (index < s.Length && s[index] != '>') {
                raw.Append(s[index]);
                index++;
            }

            if (index < s.Length && s[index] == '>') {
                index++;
            }

            return DecodeHexPdfString(raw.ToString());
        }

        static bool TryReadNumber(string s, ref int index, out double value) {
            int start = index;
            if (s[index] == '+' || s[index] == '-') {
                index++;
            }

            while (index < s.Length && (char.IsDigit(s[index]) || s[index] == '.')) {
                index++;
            }

#if NET8_0_OR_GREATER
            return double.TryParse(s.AsSpan(start, index - start), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value);
#else
            return double.TryParse(s.Substring(start, index - start), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value);
#endif
        }

        static bool IsNumberStart(char ch) => ch == '+' || ch == '-' || ch == '.' || char.IsDigit(ch);
        static double ToVisualAdvance(double tjAdjustment, double currentFontSize, double currentHorizontalScale) => -tjAdjustment / 1000.0 * currentFontSize * currentHorizontalScale;
    }

    private static bool TryAppendNextLineShowText(string line, Action<string> appendText, Action requestSpace) {
        Match literalQuote = QuoteLiteralRegex.Match(line);
        if (literalQuote.Success) {
            requestSpace();
            appendText(UnescapePdfLiteral(literalQuote.Groups["txt"].Value));
            return true;
        }

        Match hexQuote = QuoteHexRegex.Match(line);
        if (hexQuote.Success) {
            requestSpace();
            appendText(DecodeHexPdfString(hexQuote.Groups["txt"].Value));
            return true;
        }

        Match literalDoubleQuote = DoubleQuoteLiteralRegex.Match(line);
        if (literalDoubleQuote.Success) {
            requestSpace();
            appendText(UnescapePdfLiteral(literalDoubleQuote.Groups["txt"].Value));
            return true;
        }

        Match hexDoubleQuote = DoubleQuoteHexRegex.Match(line);
        if (hexDoubleQuote.Success) {
            requestSpace();
            appendText(DecodeHexPdfString(hexDoubleQuote.Groups["txt"].Value));
            return true;
        }

        return false;
    }

    private static string DecodeMetadataHexString(string raw) {
        if (string.IsNullOrWhiteSpace(raw)) {
            return string.Empty;
        }

        var bytes = DecodeHexBytes(raw);
        if (bytes.Length >= 2) {
            if (bytes[0] == 0xFE && bytes[1] == 0xFF) {
                return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            }

            if (bytes[0] == 0xFF && bytes[1] == 0xFE) {
                return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            }
        }

        return PdfWinAnsiEncoding.Decode(bytes);
    }

    private static byte[] DecodeHexBytes(string s) {
        var hex = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++) {
            char ch = s[i];
            if (!char.IsWhiteSpace(ch)) hex.Append(ch);
        }

        if (hex.Length % 2 == 1) hex.Append('0');

        var bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++) {
            int hi = HexNibble(hex[i * 2]);
            int lo = HexNibble(hex[i * 2 + 1]);
            bytes[i] = (byte)((hi << 4) | lo);
        }

        return bytes;

        static int HexNibble(char c) {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
            if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
            throw new FormatException($"Invalid hex character '{c}'.");
        }
    }

    private static List<PdfDictionary> CollectPages(Dictionary<int, PdfIndirectObject> objects) {
        var result = new List<PdfDictionary>();
        int? catalogId = null;
        foreach (var kv in objects) {
            if (kv.Value.Value is PdfDictionary dict && dict.Get<PdfName>("Type")?.Name == "Catalog") {
                catalogId = kv.Key;
                break;
            }
        }

        if (catalogId is int cat &&
            objects.TryGetValue(cat, out var catalogObject) &&
            catalogObject.Value is PdfDictionary catalog &&
            ResolveDict(catalog.Items.TryGetValue("Pages", out var pagesObj) ? pagesObj : null, objects) is PdfDictionary pagesRoot) {
            TraversePagesNode(pagesRoot, objects, result, new HashSet<int>());
        }

        if (result.Count > 0) {
            return result;
        }

        foreach (var kv in objects.OrderBy(k => k.Key)) {
            if (kv.Value.Value is PdfDictionary dict && IsLeafPage(dict, objects)) {
                result.Add(dict);
            }
        }

        return result;
    }

    private static void TraversePagesNode(
        PdfDictionary node,
        Dictionary<int, PdfIndirectObject> objects,
        List<PdfDictionary> result,
        HashSet<int> visited) {
        string? type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page" || (type is null && IsLeafPage(node, objects))) {
            int objectNumber = FindObjectNumberFor(node, objects);
            if (objectNumber <= 0 || visited.Add(objectNumber)) {
                result.Add(node);
            }
            return;
        }

        var kids = ResolveArray(node.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null, objects);
        if (kids is null) {
            return;
        }

        foreach (var kid in kids.Items) {
            var child = ResolveDict(kid, objects);
            if (child is not null) {
                TraversePagesNode(child, objects, result, visited);
            }
        }
    }

    private static bool IsLeafPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        if (ResolveArray(page.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null, objects) is not null) {
            return false;
        }

        if (!page.Items.ContainsKey("Contents")) {
            return false;
        }

        string? type = page.Get<PdfName>("Type")?.Name;
        if (type == "Page") {
            return true;
        }

        return type is null &&
               (page.Items.ContainsKey("Resources") || GetInheritedValue(page, "Resources", objects) is not null) &&
               (page.Items.ContainsKey("MediaBox") ||
                page.Items.ContainsKey("CropBox") ||
                GetInheritedValue(page, "MediaBox", objects) is not null ||
                GetInheritedValue(page, "CropBox", objects) is not null);
    }

    private static List<int> GetContentIds(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var ids = new List<int>();
        if (!page.Items.TryGetValue("Contents", out var contents)) {
            return ids;
        }

        if (contents is PdfReference reference) {
            if (objects.TryGetValue(reference.ObjectNumber, out var indirect) &&
                indirect.Value is PdfArray referencedArray) {
                AppendContentIds(referencedArray, ids);
            } else {
                ids.Add(reference.ObjectNumber);
            }
            return ids;
        }

        if (contents is PdfArray arr) {
            AppendContentIds(arr, ids);
        }

        return ids;
    }

    private static PdfObject? GetInheritedValue(PdfDictionary start, string key, Dictionary<int, PdfIndirectObject> objects) {
        PdfDictionary? current = start;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue(key, out var value)) {
                return value;
            }

            if (!current.Items.TryGetValue("Parent", out var parentObj)) {
                break;
            }

            current = ResolveDict(parentObj, objects);
        }

        return null;
    }

    private static PdfDictionary? ResolveDict(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfDictionary dict) {
            return dict;
        }

        if (obj is PdfReference reference &&
            objects.TryGetValue(reference.ObjectNumber, out var indirect) &&
            indirect.Value is PdfDictionary referencedDict) {
            return referencedDict;
        }

        return null;
    }

    private static PdfArray? ResolveArray(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfArray arr) {
            return arr;
        }

        if (obj is PdfReference reference &&
            objects.TryGetValue(reference.ObjectNumber, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            return referencedArray;
        }

        return null;
    }

    private static void AppendContentIds(PdfArray contentArray, List<int> ids) {
        foreach (var item in contentArray.Items) {
            if (item is PdfReference itemReference) {
                ids.Add(itemReference.ObjectNumber);
            }
        }
    }

    private static int FindObjectNumberFor(PdfDictionary dict, Dictionary<int, PdfIndirectObject> objects) {
        foreach (var kv in objects) {
            if (ReferenceEquals(kv.Value.Value, dict)) {
                return kv.Key;
            }
        }

        return 0;
    }
}
