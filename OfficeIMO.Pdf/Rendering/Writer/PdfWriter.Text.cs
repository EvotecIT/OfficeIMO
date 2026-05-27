using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private const double DefaultParagraphTabStopWidth = 36D;
    private static readonly char[] TokenSplitChars = new[] { ' ', '\n', '\t' };
    private static readonly char[] HardLineSplitChars = new[] { '\n' };
    private static readonly char[] SoftLineSplitChars = new[] { ' ', '\t' };
    private static string EscapeText(string s) => PdfSyntaxEscaper.EscapeLiteralContent(s);

    private static string EncodeWinAnsiHex(string s) {
        var bytes = PdfWinAnsiEncoding.Encode(s);
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        return sb.ToString();
    }

    private static System.Collections.Generic.List<string> WrapMonospace(string text, double widthPts, double fontSize, double glyphWidthEm) {
        double glyphWidth = fontSize * glyphWidthEm;
        int maxChars = Math.Max(8, (int)Math.Floor(widthPts / glyphWidth));
        int maxUnspacedTokenChars = Math.Max(1, (int)Math.Floor(widthPts / (fontSize * Math.Max(glyphWidthEm, 0.9))));
        var hardLines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split(HardLineSplitChars, StringSplitOptions.None);
        var lines = new System.Collections.Generic.List<string>();
        var line = new StringBuilder();
        void AddWrappedWord(string word) {
            if (word.Length <= maxChars) {
                line.Append(word);
                return;
            }

            for (int i = 0; i < word.Length; i += maxUnspacedTokenChars) {
                var chunk = word.Substring(i, Math.Min(maxUnspacedTokenChars, word.Length - i));
                if (chunk.Length == maxUnspacedTokenChars) {
                    lines.Add(chunk);
                } else {
                    line.Append(chunk);
                }
            }
        }

        void AddSoftWrappedLine(string hardLine) {
            int startingLineCount = lines.Count;
            var words = hardLine.Split(SoftLineSplitChars, StringSplitOptions.None);
            foreach (var w in words) {
                if (line.Length == 0) {
                    AddWrappedWord(w);
                } else {
                    if (line.Length + 1 + w.Length <= maxChars) {
                        line.Append(' ').Append(w);
                    } else {
                        lines.Add(line.ToString());
                        line.Clear();
                        AddWrappedWord(w);
                    }
                }
            }

            if (line.Length > 0) {
                lines.Add(line.ToString());
                line.Clear();
            } else if (hardLine.Length == 0 && lines.Count == startingLineCount) {
                lines.Add(string.Empty);
            }
        }

        for (int i = 0; i < hardLines.Length; i++) {
            AddSoftWrappedLine(hardLines[i]);
        }

        if (lines.Count == 0) lines.Add(string.Empty);
        return lines;
    }

    private static System.Collections.Generic.List<string> WrapSimpleText(string text, double widthPts, PdfStandardFont font, double fontSize) {
        var hardLines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split(HardLineSplitChars, StringSplitOptions.None);
        var lines = new System.Collections.Generic.List<string>();
        double maxWidth = Math.Max(1D, widthPts);
        double spaceWidth = EstimateSimpleTextWidth(" ", font, fontSize);

        void FlushLine(StringBuilder current, ref double currentWidth) {
            if (current.Length > 0) {
                lines.Add(current.ToString());
                current.Clear();
                currentWidth = 0D;
            }
        }

        void AppendLongToken(string token, StringBuilder current, ref double currentWidth) {
            FlushLine(current, ref currentWidth);
            for (int i = 0; i < token.Length; i++) {
                string character = token.Substring(i, 1);
                double characterWidth = EstimateSimpleTextWidth(character, font, fontSize);
                if (current.Length > 0 && currentWidth + characterWidth > maxWidth) {
                    FlushLine(current, ref currentWidth);
                }

                current.Append(character);
                currentWidth += characterWidth;
            }
        }

        for (int hardLineIndex = 0; hardLineIndex < hardLines.Length; hardLineIndex++) {
            string hardLine = hardLines[hardLineIndex];
            int startingLineCount = lines.Count;
            var current = new StringBuilder();
            double currentWidth = 0D;
            bool pendingSpace = false;
            int index = 0;

            while (index < hardLine.Length) {
                int nextWhitespace = hardLine.IndexOfAny(SoftLineSplitChars, index);
                string token;
                if (nextWhitespace == -1) {
                    token = hardLine.Substring(index);
                    index = hardLine.Length;
                } else {
                    token = hardLine.Substring(index, nextWhitespace - index);
                    index = nextWhitespace + 1;
                }

                if (token.Length > 0) {
                    double tokenWidth = EstimateSimpleTextWidth(token, font, fontSize);
                    if (tokenWidth > maxWidth) {
                        AppendLongToken(token, current, ref currentWidth);
                    } else {
                        double neededWidth = current.Length == 0 ? tokenWidth : (pendingSpace ? spaceWidth : 0D) + tokenWidth;
                        if (current.Length > 0 && currentWidth + neededWidth > maxWidth) {
                            FlushLine(current, ref currentWidth);
                        }

                        if (current.Length > 0 && pendingSpace) {
                            current.Append(' ');
                            currentWidth += spaceWidth;
                        }

                        current.Append(token);
                        currentWidth += tokenWidth;
                    }

                    pendingSpace = false;
                }

                if (nextWhitespace != -1) {
                    pendingSpace = true;
                }
            }

            FlushLine(current, ref currentWidth);
            if (hardLine.Length == 0 && lines.Count == startingLineCount) {
                lines.Add(string.Empty);
            }
        }

        if (lines.Count == 0) lines.Add(string.Empty);
        return lines;
    }

    // Rich paragraph layout
    private sealed record RichSeg(string Text, bool Bold, bool Italic, bool Underline, bool Strike, PdfColor? Color, string? Uri, string? DestinationName, string? Contents, PdfStandardFont Font, PdfTextBaseline Baseline, bool LeadingSpace = false, double LeadingAdvance = 0, bool LeadingSpaceIsExpandable = true, PdfTabLeaderStyle LeadingTabLeader = PdfTabLeaderStyle.None, bool EndsWithHardBreak = false);

    private static double MeasureRichText(string text, PdfStandardFont font, double fontSize) =>
        EstimateSimpleTextWidth(text, font, fontSize);

    private static double EffectiveRichFontSize(double fontSize, PdfTextBaseline baseline) =>
        baseline == PdfTextBaseline.Normal ? fontSize : fontSize * 0.65;

    private static double TextRiseForBaseline(double fontSize, PdfTextBaseline baseline) => baseline switch {
        PdfTextBaseline.Superscript => fontSize * 0.35,
        PdfTextBaseline.Subscript => -fontSize * 0.18,
        _ => 0
    };

    private static double MeasureRichText(string text, PdfStandardFont font, double fontSize, PdfTextBaseline baseline) =>
        EstimateSimpleTextWidth(text, font, EffectiveRichFontSize(fontSize, baseline));

    private static double CalculateDefaultTabAdvance(double lineWidth, double spaceWidth, double tabStopWidth = DefaultParagraphTabStopWidth) {
        if (lineWidth < 0 || double.IsNaN(lineWidth) || double.IsInfinity(lineWidth) ||
            tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth)) {
            return spaceWidth;
        }

        double nextStop = (Math.Floor(lineWidth / tabStopWidth) + 1D) * tabStopWidth;
        return Math.Max(spaceWidth, nextStop - lineWidth);
    }

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRuns(System.Collections.Generic.IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts = null, double tabStopWidth = DefaultParagraphTabStopWidth) {
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new() };
        var heights = new System.Collections.Generic.List<double>();
        double lineWidth = 0;
        double pendingLeadingAdvance = 0;
        bool pendingLeadingIsExpandable = true;
        PdfTabLeaderStyle pendingLeadingTabLeader = PdfTabLeaderStyle.None;
        double CurrentMaxWidth() => lines.Count == 1 ? firstLineWidthPts ?? maxWidthPts : maxWidthPts;

        void MarkCurrentLineHardBreak() {
            var currentLine = lines[lines.Count - 1];
            if (currentLine.Count == 0) {
                return;
            }

            var lastSegment = currentLine[currentLine.Count - 1];
            currentLine[currentLine.Count - 1] = lastSegment with { EndsWithHardBreak = true };
        }

        foreach (var run in runs) {
            string text = (run.Text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            bool bold = run.Bold;
            bool underline = run.Underline;
            bool strike = run.Strike;
            bool italic = run.Italic;
            var color = run.Color;
            string? uri = run.LinkUri;
            string? destinationName = run.LinkDestinationName;
            string? contents = run.LinkContents;
            var baseline = run.Baseline;
            var tabLeader = run.TabLeader;
            var fontForRun = (bold && italic) ? ChooseBoldItalic(baseFont) : bold ? ChooseBold(baseFont) : italic ? ChooseItalic(baseFont) : baseFont;
            double spaceW = MeasureRichText(" ", fontForRun, fontSize, baseline);
            int idx = 0;
            while (idx < text.Length) {
                int nextWs = text.IndexOfAny(TokenSplitChars, idx);
                bool hadNewline = false;
                string token;
                if (nextWs == -1) { token = text.Substring(idx); idx = text.Length; } else {
                    token = text.Substring(idx, nextWs - idx);
                    hadNewline = text[nextWs] == '\n';
                    idx = nextWs + 1;
                }
                double tokenW = MeasureRichText(token, fontForRun, fontSize, baseline);
                var lastLine = lines[lines.Count - 1];
                double needed = lastLine.Count == 0 ? tokenW : pendingLeadingAdvance + tokenW;
                double currentMaxWidth = CurrentMaxWidth();

                if (tokenW > currentMaxWidth) {
                    if (lastLine.Count > 0) { heights.Add(lineHeight); lines.Add(new()); lineWidth = 0; lastLine = lines[lines.Count - 1]; }
                    pendingLeadingAdvance = 0;
                    pendingLeadingIsExpandable = true;
                    pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                    int pos = 0;
                    while (pos < token.Length) {
                        int take = 0;
                        double chunkW = 0;
                        currentMaxWidth = CurrentMaxWidth();
                        while (pos + take < token.Length) {
                            double charW = MeasureRichText(token.Substring(pos + take, 1), fontForRun, fontSize, baseline);
                            if (take > 0 && chunkW + charW > currentMaxWidth) {
                                break;
                            }

                            chunkW += charW;
                            take++;
                            if (chunkW >= currentMaxWidth) {
                                break;
                            }
                        }

                        if (take == 0) {
                            take = 1;
                            chunkW = MeasureRichText(token.Substring(pos, 1), fontForRun, fontSize, baseline);
                        }

                        string chunk = token.Substring(pos, take);
                        lastLine.Add(new RichSeg(chunk, bold, italic, underline, strike, color, uri, destinationName, contents, fontForRun, baseline));
                        lineWidth += chunkW;
                        pos += take;
                        if (pos < token.Length) { heights.Add(lineHeight); lines.Add(new()); lineWidth = 0; lastLine = lines[lines.Count - 1]; }
                    }
                    if (hadNewline) {
                        MarkCurrentLineHardBreak();
                        heights.Add(lineHeight);
                        lines.Add(new());
                        lineWidth = 0;
                        pendingLeadingAdvance = 0;
                        pendingLeadingIsExpandable = true;
                        pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                    } else if (nextWs != -1) {
                        bool hadTab = text[nextWs] == '\t';
                        pendingLeadingAdvance = hadTab ? CalculateDefaultTabAdvance(lineWidth, spaceW, tabStopWidth) : spaceW;
                        pendingLeadingIsExpandable = !hadTab;
                        pendingLeadingTabLeader = hadTab ? tabLeader : PdfTabLeaderStyle.None;
                    }
                    continue;
                }
                needed = lastLine.Count == 0 ? tokenW : pendingLeadingAdvance + tokenW;
                if (lineWidth + needed > currentMaxWidth && lastLine.Count > 0) {
                    heights.Add(lineHeight);
                    lines.Add(new());
                    lineWidth = 0;
                }
                if (token.Length > 0) {
                    bool needsLeadingSpace = lineWidth > 0 && pendingLeadingAdvance > 0;
                    double leadingAdvance = needsLeadingSpace ? pendingLeadingAdvance : 0;
                    double segmentWidth = tokenW + leadingAdvance;
                    var segmentLeader = needsLeadingSpace ? pendingLeadingTabLeader : PdfTabLeaderStyle.None;
                    lines[lines.Count - 1].Add(new RichSeg(token, bold, italic, underline, strike, color, uri, destinationName, contents, fontForRun, baseline, needsLeadingSpace, leadingAdvance, pendingLeadingIsExpandable, segmentLeader));
                    lineWidth += segmentWidth;
                    pendingLeadingAdvance = 0;
                    pendingLeadingIsExpandable = true;
                    pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                }
                if (hadNewline) {
                    MarkCurrentLineHardBreak();
                    heights.Add(lineHeight);
                    lines.Add(new());
                    lineWidth = 0;
                    pendingLeadingAdvance = 0;
                    pendingLeadingIsExpandable = true;
                    pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                } else if (nextWs != -1) {
                    bool hadTab = text[nextWs] == '\t';
                    pendingLeadingAdvance = hadTab ? CalculateDefaultTabAdvance(lineWidth, spaceW, tabStopWidth) : spaceW;
                    pendingLeadingIsExpandable = !hadTab;
                    pendingLeadingTabLeader = hadTab ? tabLeader : PdfTabLeaderStyle.None;
                }
            }
        }
        if (lines.Count > 0 && lines[lines.Count - 1].Count == 0) { lines.RemoveAt(lines.Count - 1); }
        if (heights.Count < lines.Count) heights.Add(lineHeight);
        return (lines, heights);
    }

    private static void WriteRichParagraph(StringBuilder sb, RichParagraphBlock block, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, PdfOptions opts, double startY, double fontSize, double defaultLeading, System.Collections.Generic.List<LinkAnnotation> annots, double? xOverride = null, double? widthOverride = null, double? firstLineXOverride = null, double? firstLineWidthOverride = null) {
        double widthContent = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double widthUsed = widthOverride ?? widthContent;
        var underlines = new System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color)>();
        var strikes = new System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color)>();

        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .TextLeading(defaultLeading);
        double xOrigin = xOverride ?? opts.MarginLeft;

        for (int li = 0; li < lines.Count; li++) {
            double lineWidthUsed = li == 0 ? firstLineWidthOverride ?? widthUsed : widthUsed;
            double lineXOrigin = li == 0 ? firstLineXOverride ?? xOrigin : xOrigin;
            var segs = lines[li];
            int segCount = segs.Count;
            double[] segWidths = segCount > 0 ? new double[segCount] : System.Array.Empty<double>();
            double baseLineW = 0;
            int gapsCount = 0;
            for (int si = 0; si < segCount; si++) {
                var seg = segs[si];
                double w = MeasureRichText(seg.Text, seg.Font, fontSize, seg.Baseline);
                if (seg.LeadingSpace) {
                    w += seg.LeadingAdvance > 0 ? seg.LeadingAdvance : MeasureRichText(" ", seg.Font, fontSize, seg.Baseline);
                    if (seg.LeadingSpaceIsExpandable) {
                        gapsCount++;
                    }
                }
                segWidths[si] = w;
                baseLineW += w;
            }
            bool lineEndsWithHardBreak = segs.Any(seg => seg.EndsWithHardBreak);
            bool justify = block.Align == PdfAlign.Justify && !lineEndsWithHardBreak && li != lines.Count - 1 && gapsCount > 0 && lineWidthUsed > baseLineW;
            double wordSpacing = justify ? (lineWidthUsed - baseLineW) / gapsCount : 0;

            double lineWForAlign = justify ? lineWidthUsed : baseLineW;
            double dx = 0;
            if (block.Align == PdfAlign.Center) dx = Math.Max(0, (lineWidthUsed - lineWForAlign) / 2);
            else if (block.Align == PdfAlign.Right) dx = Math.Max(0, lineWidthUsed - lineWForAlign);
            content
                .TextMatrix(lineXOrigin + dx, startY - li * defaultLeading)
                .WordSpacing(wordSpacing);

            double xCursor = dx;
            double currentTextRise = 0;
            for (int si = 0; si < segs.Count; si++) {
                var s = segs[si];
                string fontRes = (s.Bold && s.Italic) ? "F4" : s.Bold ? "F2" : s.Italic ? "F3" : "F1";
                double runFontSize = EffectiveRichFontSize(fontSize, s.Baseline);
                double textRise = TextRiseForBaseline(fontSize, s.Baseline);
                content.Font(fontRes, runFontSize);
                if (Math.Abs(textRise - currentTextRise) > 0.0001) {
                    content.TextRise(textRise);
                    currentTextRise = textRise;
                }

                var color = s.Color ?? block.DefaultColor ?? opts.DefaultTextColor;
                if (color.HasValue) content.FillColor(color.Value);
                if (s.LeadingSpace) {
                    double baseGap = s.LeadingAdvance > 0 ? s.LeadingAdvance : MeasureRichText(" ", s.Font, fontSize, s.Baseline);
                    double gap = baseGap + (s.LeadingSpaceIsExpandable ? wordSpacing : 0);
                    double visibleGap = 0;
                    if (!s.LeadingSpaceIsExpandable && Math.Abs(wordSpacing) > 0.0001) {
                        content.WordSpacing(0);
                    }

                    if (s.LeadingTabLeader == PdfTabLeaderStyle.Dots) {
                        string leader = BuildDotLeaderText(gap, s.Font, fontSize, s.Baseline);
                        if (leader.Length > 0) {
                            content.ShowHexText(EncodeWinAnsiHex(leader));
                            visibleGap = MeasureRichText(leader, s.Font, fontSize, s.Baseline);
                        }
                    } else {
                        content.ShowHexText("20");
                        visibleGap = MeasureRichText(" ", s.Font, fontSize, s.Baseline) + (s.LeadingSpaceIsExpandable ? wordSpacing : 0);
                    }

                    if (!s.LeadingSpaceIsExpandable && Math.Abs(wordSpacing) > 0.0001) {
                        content.WordSpacing(wordSpacing);
                    }

                    double extraAdvance = gap - visibleGap;
                    if (extraAdvance > 0.0001) {
                        content.MoveText(extraAdvance, 0);
                    }

                    xCursor += gap;
                }
                double segmentStartX = xCursor;
                content.ShowHexText(EncodeWinAnsiHex(s.Text));
                double wSeg = MeasureRichText(s.Text, s.Font, fontSize, s.Baseline);
                double baselineY = startY - li * defaultLeading + textRise;

                if (s.Underline) {
                    var ulColor = (s.Color ?? block.DefaultColor ?? opts.DefaultTextColor) ?? PdfColor.Black;
                    double yLine = baselineY - runFontSize * 0.15;
                    underlines.Add((lineXOrigin + segmentStartX, lineXOrigin + segmentStartX + wSeg, yLine, ulColor));
                }
                if (s.Strike) {
                    var stColor = (s.Color ?? block.DefaultColor ?? opts.DefaultTextColor) ?? PdfColor.Black;
                    double yLine = baselineY + runFontSize * 0.32;
                    strikes.Add((lineXOrigin + segmentStartX, lineXOrigin + segmentStartX + wSeg, yLine, stColor));
                }
                if (!string.IsNullOrEmpty(s.Uri) || !string.IsNullOrEmpty(s.DestinationName)) {
                    var fontForMetrics = s.Font;
                    double asc = GetAscender(fontForMetrics, runFontSize);
                    double desc = GetDescender(fontForMetrics, runFontSize);
                    double x1 = lineXOrigin + segmentStartX;
                    double x2 = x1 + wSeg;
                    double y1 = baselineY - desc;
                    double y2 = baselineY + asc;
                    annots.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = s.Uri, DestinationName = s.DestinationName, Contents = s.Contents });
                }
                xCursor += wSeg;
            }
            if (Math.Abs(currentTextRise) > 0.0001) {
                content.TextRise(0);
            }
        }
        content
            .WordSpacing(0)
            .EndText();

        foreach (var ul in underlines) {
            new ContentStreamBuilder(sb)
                .SaveState()
                .StrokeColor(ul.Color)
                .LineWidth(0.5)
                .MoveTo(ul.X1, ul.Y)
                .LineTo(ul.X2, ul.Y)
                .StrokePath()
                .RestoreState();
        }
        foreach (var st in strikes) {
            new ContentStreamBuilder(sb)
                .SaveState()
                .StrokeColor(st.Color)
                .LineWidth(0.5)
                .MoveTo(st.X1, st.Y)
                .LineTo(st.X2, st.Y)
                .StrokePath()
                .RestoreState();
        }
    }

    private static string BuildDotLeaderText(double gap, PdfStandardFont font, double fontSize, PdfTextBaseline baseline) {
        double dotWidth = MeasureRichText(".", font, fontSize, baseline);
        if (dotWidth <= 0 || gap <= dotWidth * 3D) {
            return string.Empty;
        }

        int count = Math.Max(3, (int)Math.Floor(gap / dotWidth));
        return new string('.', count);
    }
}
