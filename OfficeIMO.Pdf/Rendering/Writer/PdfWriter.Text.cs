using System.Text;
using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly char[] TokenSplitChars = new[] { ' ', '\n' };
    private static string EscapeText(string s) => EscapeLiteral(s);

    private static string EscapeLiteral(string s) {
        if (string.IsNullOrEmpty(s)) return string.Empty;
        var sb = new StringBuilder(s.Length + 8);
        for (int i = 0; i < s.Length; i++) {
            char ch = s[i];
            switch (ch) {
                case '\\': sb.Append("\\\\"); break;
                case '(': sb.Append("\\("); break;
                case ')': sb.Append("\\)"); break;
                case '\r': sb.Append("\\r"); break;
                case '\n': sb.Append("\\n"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    if (ch < 32 || ch == 127) {
                        int v = ch;
                        sb.Append('\\')
                          .Append(((v >> 6) & 0x7).ToString(CultureInfo.InvariantCulture))
                          .Append(((v >> 3) & 0x7).ToString(CultureInfo.InvariantCulture))
                          .Append((v & 0x7).ToString(CultureInfo.InvariantCulture));
                    } else {
                        sb.Append(ch);
                    }
                    break;
            }
        }
        return sb.ToString();
    }

    private static string EncodeWinAnsiHex(string s) {
        var bytes = PdfWinAnsiEncoding.Encode(s);
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        return sb.ToString();
    }

    private static System.Collections.Generic.List<string> WrapMonospace(string text, double widthPts, double fontSize, double glyphWidthEm) {
        double glyphWidth = fontSize * glyphWidthEm;
        int maxChars = Math.Max(8, (int)Math.Floor(widthPts / glyphWidth));
        var words = text.Replace("\r", "").Split(WordSplitChars, StringSplitOptions.None);
        var lines = new System.Collections.Generic.List<string>();
        var line = new StringBuilder();
        foreach (var w in words) {
            if (w.Contains('\n')) {
                // unexpected due to split, ignore
            }
            if (line.Length == 0) {
                if (w.Length <= maxChars) line.Append(w);
                else {
                    for (int i = 0; i < w.Length; i += maxChars) {
                        var chunk = w.Substring(i, Math.Min(maxChars, w.Length - i));
                        lines.Add(chunk);
                    }
                }
            } else {
                if (line.Length + 1 + w.Length <= maxChars) {
                    line.Append(' ').Append(w);
                } else {
                    lines.Add(line.ToString());
                    line.Clear();
                    line.Append(w);
                }
            }
        }
        if (line.Length > 0) lines.Add(line.ToString());
        if (lines.Count == 0) lines.Add(string.Empty);
        return lines;
    }

    // Rich paragraph layout
    private sealed record RichSeg(string Text, bool Bold, bool Italic, bool Underline, bool Strike, PdfColor? Color, string? Uri, PdfStandardFont Font);

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRuns(System.Collections.Generic.IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont) {
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new() };
        var heights = new System.Collections.Generic.List<double>();
        double lineWidth = 0;
        foreach (var run in runs) {
            string text = run.Text ?? string.Empty;
            bool bold = run.Bold;
            bool underline = run.Underline;
            bool strike = run.Strike;
            bool italic = run.Italic;
            var color = run.Color;
            string? uri = run.LinkUri;
            var fontForRun = (bold && italic) ? ChooseBoldItalic(baseFont) : bold ? ChooseBold(baseFont) : italic ? ChooseItalic(baseFont) : baseFont;
            double em = GlyphWidthEmFor(fontForRun);
            double spaceW = fontSize * em;
            int maxChars = System.Math.Max(1, (int)System.Math.Floor(maxWidthPts / (fontSize * em)));
            int idx = 0;
            while (idx < text.Length) {
                int nextWs = text.IndexOfAny(TokenSplitChars, idx);
                bool hadNewline = false;
                string token;
                if (nextWs == -1) { token = text.Substring(idx); idx = text.Length; }
                else {
                    token = text.Substring(idx, nextWs - idx);
                    hadNewline = text[nextWs] == '\n';
                    idx = nextWs + 1;
                }
                double tokenW = token.Length * fontSize * em;
                var lastLine = lines[lines.Count - 1];
                double needed = (lastLine.Count == 0 ? tokenW : spaceW + tokenW);

                if (tokenW > maxWidthPts) {
                    if (lastLine.Count > 0) { heights.Add(fontSize * 1.4); lines.Add(new()); lineWidth = 0; lastLine = lines[lines.Count - 1]; }
                    int pos = 0;
                    while (pos < token.Length) {
                        int take = System.Math.Min(maxChars, token.Length - pos);
                        string chunk = token.Substring(pos, take);
                        lastLine.Add(new RichSeg(chunk, bold, italic, underline, strike, color, uri, fontForRun));
                        double chunkW = take * fontSize * em;
                        lineWidth += chunkW;
                        pos += take;
                        if (pos < token.Length) { heights.Add(fontSize * 1.4); lines.Add(new()); lineWidth = 0; lastLine = lines[lines.Count - 1]; }
                    }
                    if (hadNewline) { heights.Add(fontSize * 1.4); lines.Add(new()); lineWidth = 0; }
                    continue;
                }
                if (lineWidth + needed > maxWidthPts && lastLine.Count > 0) {
                    heights.Add(fontSize * 1.4);
                    lines.Add(new());
                    lineWidth = 0;
                    lastLine = lines[lines.Count - 1];
                    needed = tokenW;
                }
                if (token.Length > 0) {
                    if (lineWidth > 0) lineWidth += spaceW;
                    lines[lines.Count - 1].Add(new RichSeg(token, bold, italic, underline, strike, color, uri, fontForRun));
                    lineWidth += tokenW;
                }
                if (hadNewline) {
                    heights.Add(fontSize * 1.4);
                    lines.Add(new());
                    lineWidth = 0;
                }
            }
        }
        if (lines.Count > 0 && lines[lines.Count - 1].Count == 0) { lines.RemoveAt(lines.Count - 1); }
        if (heights.Count < lines.Count) heights.Add(fontSize * 1.4);
        return (lines, heights);
    }

    private static void WriteRichParagraph(StringBuilder sb, RichParagraphBlock block, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, PdfOptions opts, double startY, double fontSize, double defaultLeading, System.Collections.Generic.List<LinkAnnotation> annots, double? xOverride = null, double? widthOverride = null) {
        double widthContent = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double widthUsed = widthOverride ?? widthContent;
        var underlines = new System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color)>();
        var strikes = new System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color)>();

        sb.Append("BT\n");
        sb.Append(F(defaultLeading)).Append(" TL\n");
        double xOrigin = xOverride ?? opts.MarginLeft;
        sb.Append("1 0 0 1 ").Append(F(xOrigin)).Append(' ').Append(F(startY)).Append(" Tm\n");

        for (int li = 0; li < lines.Count; li++) {
            if (li != 0) sb.Append("T*\n");
            var segs = lines[li];
            int segCount = segs.Count;
            double[] segWidths = segCount > 0 ? new double[segCount] : System.Array.Empty<double>();
            double[] gapWidths = segCount > 1 ? new double[segCount - 1] : System.Array.Empty<double>();
            double baseLineW = 0;
            for (int si = 0; si < segCount; si++) {
                var seg = segs[si];
                double emSeg = GlyphWidthEmFor(seg.Font);
                double w = seg.Text.Length * fontSize * emSeg;
                segWidths[si] = w;
                baseLineW += w;
                if (si > 0) {
                    double gap = fontSize * GlyphWidthEmFor(seg.Font);
                    gapWidths[si - 1] = gap;
                    baseLineW += gap;
                }
            }
            int gapsCount = gapWidths.Length;
            bool justify = block.Align == PdfAlign.Justify && li != lines.Count - 1 && gapsCount > 0;
            if (justify && widthUsed > baseLineW) {
                double extra = (widthUsed - baseLineW) / gapsCount;
                for (int gi = 0; gi < gapWidths.Length; gi++) gapWidths[gi] += extra;
            }

            double lineWForAlign = justify ? widthUsed : baseLineW;
            double dx = 0;
            if (block.Align == PdfAlign.Center) dx = Math.Max(0, (widthUsed - lineWForAlign) / 2);
            else if (block.Align == PdfAlign.Right) dx = Math.Max(0, widthUsed - lineWForAlign);
            if (dx != 0) sb.Append(F(dx)).Append(" 0 Td\n");

            double xCursor = dx;
            for (int si = 0; si < segs.Count; si++) {
                if (si > 0) {
                    double gapAdvance = gapWidths[si - 1];
                    sb.Append(F(gapAdvance)).Append(" 0 Td\n");
                    xCursor += gapAdvance;
                }
                var s = segs[si];
                string fontRes = (s.Bold && s.Italic) ? "F4" : s.Bold ? "F2" : s.Italic ? "F3" : "F1";
                sb.Append('/').Append(fontRes).Append(' ').Append(F(fontSize)).Append(" Tf\n");
                var color = s.Color ?? block.DefaultColor ?? opts.DefaultTextColor;
                if (color.HasValue) sb.Append(SetFillColor(color.Value));
                sb.Append('<').Append(EncodeWinAnsiHex(s.Text)).Append("> Tj\n");
                double wSeg = segWidths[si];
                sb.Append(F(wSeg)).Append(" 0 Td\n");
                double segmentStartX = xCursor;

                if (s.Underline) {
                    var ulColor = (s.Color ?? block.DefaultColor ?? opts.DefaultTextColor) ?? PdfColor.Black;
                    double yLine = startY - li * defaultLeading - fontSize * 0.15;
                    underlines.Add((xOrigin + segmentStartX, xOrigin + segmentStartX + wSeg, yLine, ulColor));
                }
                if (s.Strike) {
                    var stColor = (s.Color ?? block.DefaultColor ?? opts.DefaultTextColor) ?? PdfColor.Black;
                    double yLine = startY - li * defaultLeading + fontSize * 0.32;
                    strikes.Add((xOrigin + segmentStartX, xOrigin + segmentStartX + wSeg, yLine, stColor));
                }
                if (!string.IsNullOrEmpty(s.Uri)) {
                    double baseline = startY - li * defaultLeading;
                    var fontForMetrics = s.Font;
                    double asc = GetAscender(fontForMetrics, fontSize);
                    double desc = GetDescender(fontForMetrics, fontSize);
                    double x1 = xOrigin + segmentStartX;
                    double x2 = x1 + wSeg;
                    double y1 = baseline - desc;
                    double y2 = baseline + asc;
                    annots.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = s.Uri! });
                }
                xCursor += wSeg;
            }
            if (xCursor != 0) sb.Append(F(-xCursor)).Append(" 0 Td\n");
        }
        sb.Append("ET\n");

        foreach (var ul in underlines) {
            sb.Append("q\n");
            sb.Append(SetStrokeColor(ul.Color));
            sb.Append("0.5 w\n");
            sb.Append(F(ul.X1)).Append(' ').Append(F(ul.Y)).Append(" m ").Append(F(ul.X2)).Append(' ').Append(F(ul.Y)).Append(" l S\n");
            sb.Append("Q\n");
        }
        foreach (var st in strikes) {
            sb.Append("q\n");
            sb.Append(SetStrokeColor(st.Color));
            sb.Append("0.5 w\n");
            sb.Append(F(st.X1)).Append(' ').Append(F(st.Y)).Append(" m ").Append(F(st.X2)).Append(' ').Append(F(st.Y)).Append(" l S\n");
            sb.Append("Q\n");
        }
    }
}
