using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Minimal, zero-dependency text extractor for simple PDFs produced by OfficeIMO.Pdf
/// (uncompressed streams, literal Tj strings, line breaks via T*).
/// Not a general-purpose PDF parser; designed as a pragmatic starting point.
/// </summary>
public static class PdfTextExtractor {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
#if NET8_0_OR_GREATER
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+0\s+obj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex InfoRefRegex = new Regex(@"/Info\s+(\d+)\s+0\s+R", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex PageObjRegex = new Regex(@"<<(?:.*?)/Type\s*/Page(?:.*?)/Contents\s+(\d+)\s+0\s+R(?:.*?)/?>>", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex TjRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*Tj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
#else
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+0\s+obj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex InfoRefRegex = new Regex(@"/Info\s+(\d+)\s+0\s+R", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex PageObjRegex = new Regex(@"<<(?:.|\n|\r)*?/Type\s*/Page(?:.|\n|\r)*?/Contents\s+(\d+)\s+0\s+R(?:.|\n|\r)*?>>", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex TjRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*Tj", RegexOptions.Compiled, RegexTimeout);
#endif

    /// <summary>Extracts plain text from all pages, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(string path) {
        var bytes = File.ReadAllBytes(path);
        return ExtractAllText(bytes);
    }

    /// <summary>Extracts plain text from all pages, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(byte[] pdf) {
        var map = BuildObjectMap(pdf, out _);
        var pageContents = FindPageContentIds(pdf);
        var sb = new StringBuilder();
        for (int i = 0; i < pageContents.Count; i++) {
            if (map.TryGetValue(pageContents[i], out var obj)) {
                var m = StreamRegex.Match(obj);
                if (m.Success) {
                    if (i > 0) sb.AppendLine();
                    sb.Append(ExtractTextFromContentStream(m.Groups[1].Value));
                }
            }
        }
        return sb.ToString();
    }

    /// <summary>Gets document metadata (Title/Author/Subject/Keywords) if present; null when absent.</summary>
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(byte[] pdf) {
        var map = BuildObjectMap(pdf, out var trailer);
        var m = InfoRefRegex.Match(trailer);
        if (!m.Success) return (null, null, null, null);
        int infoId = int.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        if (!map.TryGetValue(infoId, out var obj)) return (null, null, null, null);
        string? title = ExtractLiteral(obj, "/Title");
        string? author = ExtractLiteral(obj, "/Author");
        string? subject = ExtractLiteral(obj, "/Subject");
        string? keywords = ExtractLiteral(obj, "/Keywords");
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

    private static List<int> FindPageContentIds(byte[] pdf) {
        string text = PdfEncoding.Latin1GetString(pdf);
        var ids = new List<int>();
        foreach (Match m in PageObjRegex.Matches(text)) {
            if (int.TryParse(m.Groups[1].Value, out int id)) ids.Add(id);
        }
        return ids;
    }

    private static string ExtractTextFromContentStream(string content) {
        var sb = new StringBuilder();
        bool inText = false;
        using StringReader sr = new StringReader(content);
        string? line;
        while ((line = sr.ReadLine()) is not null) {
            if (line.Contains(" BT")) { inText = true; continue; }
            if (line.Contains(" ET")) { inText = false; continue; }
            if (!inText) continue;

            // Handle T* as newline
            if (line.Trim() == "T*") { sb.AppendLine(); continue; }

            foreach (Match tj in TjRegex.Matches(line)) {
                var raw = tj.Groups["txt"].Value;
                sb.Append(UnescapePdfLiteral(raw));
            }
        }
        return sb.ToString();
    }

    internal static string UnescapePdfLiteral(string s) {
        var sb = new StringBuilder();
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c == '\\' && i + 1 < s.Length) {
                char n = s[++i];
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

    private static string? ExtractLiteral(string obj, string key) {
        int idx = obj.IndexOf(key, StringComparison.Ordinal);
        if (idx < 0) return null;
        int open = obj.IndexOf('(', idx);
        if (open < 0) return null;
        int close = FindCloseParen(obj, open);
        if (close < 0) return null;
        string raw = obj.Substring(open + 1, close - open - 1);
        return UnescapePdfLiteral(raw);
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
}
