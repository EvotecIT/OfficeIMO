using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static class TextContentParser {
    private static readonly Regex TfRe = new Regex(@"/(?<f>\w+)\s+(?<s>-?\d+(?:\.\d+)?)\s+Tf", RegexOptions.Compiled);
    private static readonly Regex TmRe = new Regex(@"(?<a>-?\d+(?:\.\d+)?)\s+(?<b>-?\d+(?:\.\d+)?)\s+(?<c>-?\d+(?:\.\d+)?)\s+(?<d>-?\d+(?:\.\d+)?)\s+(?<e>-?\d+(?:\.\d+)?)\s+(?<f>-?\d+(?:\.\d+)?)\s+Tm", RegexOptions.Compiled);
    private static readonly Regex TdRe = new Regex(@"(?<tx>-?\d+(?:\.\d+)?)\s+(?<ty>-?\d+(?:\.\d+)?)\s+Td", RegexOptions.Compiled);
    private static readonly Regex TLRe = new Regex(@"(?<lead>-?\d+(?:\.\d+)?)\s+TL", RegexOptions.Compiled);
    private static readonly Regex TjRe = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*Tj", RegexOptions.Compiled);
    private static readonly Regex TJRe = new Regex(@"\[(?<arr>[\s\S]*?)\]\s*TJ", RegexOptions.Compiled);

    public static List<PdfTextSpan> Parse(string content, System.Func<string, byte[], string> decodeWithFont) {
        var spans = new List<PdfTextSpan>();
        bool inText = false;
        string font = "F1"; double size = 12; double x = 0, y = 0; double leading = size * 1.2;
        using var sr = new StringReader(content);
        string? line;
        while ((line = sr.ReadLine()) is not null) {
            var t = line.Trim();
            if (t.EndsWith(" BT", StringComparison.Ordinal) || t == "BT") { inText = true; continue; }
            if (t.EndsWith(" ET", StringComparison.Ordinal) || t == "ET") { inText = false; continue; }
            if (!inText) continue;

            var mTf = TfRe.Match(line);
            if (mTf.Success) {
                font = mTf.Groups["f"].Value;
                size = double.Parse(mTf.Groups["s"].Value, CultureInfo.InvariantCulture);
                continue;
            }
            var mTm = TmRe.Match(line);
            if (mTm.Success) {
                x = double.Parse(mTm.Groups["e"].Value, CultureInfo.InvariantCulture);
                y = double.Parse(mTm.Groups["f"].Value, CultureInfo.InvariantCulture);
                continue;
            }
            var mTL = TLRe.Match(line);
            if (mTL.Success) { leading = double.Parse(mTL.Groups["lead"].Value, CultureInfo.InvariantCulture); continue; }
            if (t == "T*") { y -= leading; continue; }
            var mTd = TdRe.Match(line);
            if (mTd.Success) { x += double.Parse(mTd.Groups["tx"].Value, CultureInfo.InvariantCulture); y += double.Parse(mTd.Groups["ty"].Value, CultureInfo.InvariantCulture); continue; }

            foreach (Match m in TjRe.Matches(line)) {
                var raw = m.Groups["txt"].Value;
                var bytes = PdfStringParser.ParseLiteralToBytes(raw);
                var text = decodeWithFont(font, bytes);
                if (!string.IsNullOrEmpty(text)) spans.Add(new PdfTextSpan(text, font, size, x, y));
            }

            var mTJ = TJRe.Match(line);
            if (mTJ.Success) {
                string arr = mTJ.Groups["arr"].Value;
                var sb = new StringBuilder();
                // naive scan: collect (...) chunks, ignore numbers
                for (int i = 0; i < arr.Length; i++) {
                    char c = arr[i];
                    if (c == '(') {
                        int start = i + 1; bool esc = false; var sbb = new StringBuilder();
                        while (++i < arr.Length) {
                            char ch = arr[i];
                            if (esc) { sbb.Append(ch); esc = false; }
                            else if (ch == '\\') esc = true;
                            else if (ch == ')') break; else sbb.Append(ch);
                        }
                        var bytes = PdfStringParser.ParseLiteralToBytes(sbb.ToString());
                        sb.Append(decodeWithFont(font, bytes));
                    }
                }
                if (sb.Length > 0) spans.Add(new PdfTextSpan(sb.ToString(), font, size, x, y));
            }
        }
        return spans;
    }
}
