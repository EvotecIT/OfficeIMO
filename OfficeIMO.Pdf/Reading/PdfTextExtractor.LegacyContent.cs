using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

public static partial class PdfTextExtractor {
    private static string ExtractTextFromPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> parsedObjects, Dictionary<int, string> rawObjects) {
        var pageText = new StringBuilder();
        var resources = ResolveDict(GetInheritedValue(page, "Resources", parsedObjects), parsedObjects);
        var activeForms = new HashSet<PdfStream>();
    
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
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return GetMetadata(File.ReadAllBytes(path));
    }
    
    /// <summary>Gets document metadata (Title/Author/Subject/Keywords) from the current position of a readable stream.</summary>
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(Stream stream) {
        return GetMetadata(ReadAllBytes(stream));
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
        HashSet<PdfStream>? activeForms = null) {
        var sb = new StringBuilder();
        bool inText = false;
        bool pendingSpace = false;
        bool hasTextInCurrentTextObject = false;
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
                    hasTextInCurrentTextObject = false;
                    args.Clear();
                    break;
                case "ET":
                    inText = false;
                    pendingSpace = false;
                    hasTextInCurrentTextObject = false;
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
                case "TD":
                    if (inText && args.Count >= 2) {
                        double advanceX = ToDouble(args[args.Count - 2]);
                        double advanceY = ToDouble(args[args.Count - 1]);
                        if (Math.Abs(advanceY) > 0.1 && hasTextInCurrentTextObject) {
                            AppendLineBreak();
                        } else if (advanceX > 0.1) {
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
                        AppendLineBreak();
                        AppendTextRun(ToText(args[args.Count - 1]));
                    }
                    args.Clear();
                    break;
                case "\"":
                    if (inText && args.Count >= 3) {
                        AppendLineBreak();
                        AppendTextRun(ToText(args[args.Count - 1]));
                    }
                    args.Clear();
                    break;
                case "Do":
                    if (resources is not null && parsedObjects is not null && rawObjects is not null && args.Count >= 1) {
                        string formText = ExtractInvokedFormText(ToName(args[args.Count - 1]), resources, parsedObjects, rawObjects, activeForms ?? new HashSet<PdfStream>());
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
            if (inText) {
                hasTextInCurrentTextObject = true;
            }
            pendingSpace = false;
        }
    
        void RequestSpace() {
            pendingSpace = true;
        }
    
        void AppendLineBreak() {
            if (sb.Length > 0) {
                sb.AppendLine();
            }
            pendingSpace = false;
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
        HashSet<PdfStream> activeForms) {
        if (!TryGetFormStream(resources, formName, parsedObjects, out var formStream)) {
            return string.Empty;
        }
    
        if (!activeForms.Add(formStream)) {
            return string.Empty;
        }
    
        try {
            string content = DecodeStreamContent(formStream, parsedObjects);
            var formResources = ResolveDict(formStream.Dictionary.Items.TryGetValue("Resources", out var resObj) ? resObj : null, parsedObjects) ?? resources;
            return ExtractTextFromContentStream(content, formResources, parsedObjects, rawObjects, activeForms);
        } finally {
            activeForms.Remove(formStream);
        }
    }
    
    private static bool TryGetFormStream(
        PdfDictionary resources,
        string name,
        Dictionary<int, PdfIndirectObject> parsedObjects,
        out PdfStream formStream) {
        formStream = null!;
    
        if (!resources.Items.TryGetValue("XObject", out var xObjectObj)) {
            return false;
        }
    
        var xObjectDict = ResolveDict(xObjectObj, parsedObjects);
        if (xObjectDict is null || !xObjectDict.Items.TryGetValue(name, out var formObj)) {
            return false;
        }
    
        if (formObj is PdfReference formRef &&
            PdfObjectLookup.TryGet(parsedObjects, formRef, out var indirectForm) &&
            indirectForm.Value is PdfStream referencedStream &&
            string.Equals(referencedStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
            formStream = referencedStream;
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
}
