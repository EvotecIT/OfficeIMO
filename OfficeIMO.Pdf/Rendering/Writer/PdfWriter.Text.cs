using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private const double DefaultParagraphTabStopWidth = 36D;
    private static readonly char[] TokenSplitChars = new[] { ' ', '\n', '\t' };
    private static readonly char[] HardLineSplitChars = new[] { '\n' };
    private static readonly char[] SoftLineSplitChars = new[] { ' ', '\t' };
    private static readonly char[] DecimalTabAnchorChars = new[] { '.', ',' };
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
    private sealed class RichSeg {
        public RichSeg(
            string text,
            bool bold,
            bool italic,
            bool underline,
            bool strike,
            PdfColor? color,
            PdfColor? backgroundColor,
            string? uri,
            string? destinationName,
            string? contents,
            PdfStandardFont font,
            double fontSize,
            PdfTextBaseline baseline,
            bool leadingSpace = false,
            double leadingAdvance = 0,
            bool leadingSpaceIsExpandable = true,
            PdfTabLeaderStyle leadingTabLeader = PdfTabLeaderStyle.None,
            bool endsWithHardBreak = false) {
            Text = text;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Strike = strike;
            Color = color;
            BackgroundColor = backgroundColor;
            Uri = uri;
            DestinationName = destinationName;
            Contents = contents;
            Font = font;
            FontSize = fontSize;
            Baseline = baseline;
            LeadingSpace = leadingSpace;
            LeadingAdvance = leadingAdvance;
            LeadingSpaceIsExpandable = leadingSpaceIsExpandable;
            LeadingTabLeader = leadingTabLeader;
            EndsWithHardBreak = endsWithHardBreak;
        }

        public string Text { get; }

        public bool Bold { get; }

        public bool Italic { get; }

        public bool Underline { get; }

        public bool Strike { get; }

        public PdfColor? Color { get; }

        public PdfColor? BackgroundColor { get; }

        public string? Uri { get; }

        public string? DestinationName { get; }

        public string? Contents { get; }

        public PdfStandardFont Font { get; }

        public double FontSize { get; }

        public PdfTextBaseline Baseline { get; }

        public bool LeadingSpace { get; }

        public double LeadingAdvance { get; }

        public bool LeadingSpaceIsExpandable { get; }

        public PdfTabLeaderStyle LeadingTabLeader { get; }

        public bool EndsWithHardBreak { get; }

        public RichSeg WithEndsWithHardBreak() =>
            new RichSeg(Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, Uri, DestinationName, Contents, Font, FontSize, Baseline, LeadingSpace, LeadingAdvance, LeadingSpaceIsExpandable, LeadingTabLeader, true);

        public RichSeg WithoutLink() =>
            new RichSeg(Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, null, null, null, Font, FontSize, Baseline, LeadingSpace, LeadingAdvance, LeadingSpaceIsExpandable, LeadingTabLeader, EndsWithHardBreak);
    }

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

    private static double MeasureRichLineWidth(System.Collections.Generic.IReadOnlyList<RichSeg> line) {
        double width = 0D;
        for (int index = 0; index < line.Count; index++) {
            RichSeg segment = line[index];
            if (segment.LeadingSpace) {
                width += segment.LeadingAdvance > 0
                    ? segment.LeadingAdvance
                    : MeasureRichText(" ", segment.Font, segment.FontSize, segment.Baseline);
            }

            width += MeasureRichText(segment.Text, segment.Font, segment.FontSize, segment.Baseline);
        }

        return width;
    }

    private static double CalculateDefaultTabAdvance(double lineWidth, double spaceWidth, double tabStopWidth = DefaultParagraphTabStopWidth) {
        if (lineWidth < 0 || double.IsNaN(lineWidth) || double.IsInfinity(lineWidth) ||
            tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth)) {
            return spaceWidth;
        }

        double nextStop = (Math.Floor(lineWidth / tabStopWidth) + 1D) * tabStopWidth;
        return Math.Max(spaceWidth, nextStop - lineWidth);
    }

    private static double MeasureDecimalAnchorWidth(string text, PdfStandardFont font, double fontSize, PdfTextBaseline baseline) {
        if (string.IsNullOrEmpty(text)) {
            return 0D;
        }

        int decimalIndex = text.IndexOfAny(DecimalTabAnchorChars);
        if (decimalIndex < 0) {
            return MeasureRichText(text, font, fontSize, baseline);
        }

        return MeasureRichText(text.Substring(0, decimalIndex), font, fontSize, baseline);
    }

    private static double CalculateTabAdvance(double lineWidth, double followingTextWidth, double spaceWidth, PdfTabAlignment alignment, double tabStopWidth = DefaultParagraphTabStopWidth, string followingText = "", PdfStandardFont followingFont = PdfStandardFont.Helvetica, double fontSize = 12D, PdfTextBaseline baseline = PdfTextBaseline.Normal) {
        if (alignment == PdfTabAlignment.Left) {
            return CalculateDefaultTabAdvance(lineWidth, spaceWidth, tabStopWidth);
        }

        if (lineWidth < 0 || double.IsNaN(lineWidth) || double.IsInfinity(lineWidth) ||
            followingTextWidth < 0 || double.IsNaN(followingTextWidth) || double.IsInfinity(followingTextWidth) ||
            tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth)) {
            return spaceWidth;
        }

        double anchorWidth = alignment switch {
            PdfTabAlignment.Center => followingTextWidth / 2D,
            PdfTabAlignment.Right => followingTextWidth,
            PdfTabAlignment.DecimalSeparator => MeasureDecimalAnchorWidth(followingText, followingFont, fontSize, baseline),
            _ => followingTextWidth
        };
        double nextStop = (Math.Floor(lineWidth / tabStopWidth) + 1D) * tabStopWidth;
        double advance = nextStop - anchorWidth - lineWidth;
        if (advance < spaceWidth) {
            double stopsToAdd = Math.Ceiling((spaceWidth - advance) / tabStopWidth);
            nextStop += Math.Max(1D, stopsToAdd) * tabStopWidth;
            advance = nextStop - anchorWidth - lineWidth;
        }

        return Math.Max(spaceWidth, advance);
    }

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRuns(System.Collections.Generic.IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts = null, double tabStopWidth = DefaultParagraphTabStopWidth) {
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new() };
        var heights = new System.Collections.Generic.List<double>();
        double lineWidth = 0;
        double pendingLeadingAdvance = 0;
        bool pendingLeadingIsExpandable = true;
        bool pendingLeadingIsTab = false;
        PdfTabAlignment pendingLeadingTabAlignment = PdfTabAlignment.Left;
        PdfTabLeaderStyle pendingLeadingTabLeader = PdfTabLeaderStyle.None;
        double lineHeightRatio = fontSize > 0 ? lineHeight / fontSize : 1.2D;
        double currentLineHeight = lineHeight;
        double CurrentMaxWidth() => lines.Count == 1 ? firstLineWidthPts ?? maxWidthPts : maxWidthPts;
        void RegisterLineHeight(double runFontSize) {
            currentLineHeight = Math.Max(currentLineHeight, runFontSize * lineHeightRatio);
        }

        void StartNewLine() {
            heights.Add(currentLineHeight);
            lines.Add(new());
            lineWidth = 0;
            currentLineHeight = lineHeight;
        }

        void MarkCurrentLineHardBreak() {
            var currentLine = lines[lines.Count - 1];
            if (currentLine.Count == 0) {
                return;
            }

            var lastSegment = currentLine[currentLine.Count - 1];
            currentLine[currentLine.Count - 1] = lastSegment.WithEndsWithHardBreak();
        }

        foreach (var run in runs) {
            string text = (run.Text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            bool bold = run.Bold;
            bool underline = run.Underline;
            bool strike = run.Strike;
            bool italic = run.Italic;
            var color = run.Color;
            var backgroundColor = run.BackgroundColor;
            string? uri = run.LinkUri;
            string? destinationName = run.LinkDestinationName;
            string? contents = run.LinkContents;
            var baseline = run.Baseline;
            var tabLeader = run.TabLeader;
            var tabAlignment = run.TabAlignment;
            var runBaseFont = run.Font.HasValue ? ChooseNormal(run.Font.Value) : baseFont;
            var fontForRun = (bold && italic) ? ChooseBoldItalic(runBaseFont) : bold ? ChooseBold(runBaseFont) : italic ? ChooseItalic(runBaseFont) : runBaseFont;
            double runFontSize = run.FontSize ?? fontSize;
            double spaceW = MeasureRichText(" ", fontForRun, runFontSize, baseline);
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
                double tokenW = MeasureRichText(token, fontForRun, runFontSize, baseline);
                var lastLine = lines[lines.Count - 1];
                double needed = lastLine.Count == 0 ? tokenW : pendingLeadingAdvance + tokenW;
                double currentMaxWidth = CurrentMaxWidth();

                if (tokenW > currentMaxWidth) {
                    if (lastLine.Count > 0) { StartNewLine(); lastLine = lines[lines.Count - 1]; }
                    pendingLeadingAdvance = 0;
                    pendingLeadingIsExpandable = true;
                    pendingLeadingIsTab = false;
                    pendingLeadingTabAlignment = PdfTabAlignment.Left;
                    pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                    int pos = 0;
                    while (pos < token.Length) {
                        int take = 0;
                        double chunkW = 0;
                        currentMaxWidth = CurrentMaxWidth();
                        while (pos + take < token.Length) {
                            double charW = MeasureRichText(token.Substring(pos + take, 1), fontForRun, runFontSize, baseline);
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
                            chunkW = MeasureRichText(token.Substring(pos, 1), fontForRun, runFontSize, baseline);
                        }

                        string chunk = token.Substring(pos, take);
                        lastLine.Add(new RichSeg(chunk, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline));
                        RegisterLineHeight(runFontSize);
                        lineWidth += chunkW;
                        pos += take;
                        if (pos < token.Length) { StartNewLine(); lastLine = lines[lines.Count - 1]; }
                    }
                    if (hadNewline) {
                        MarkCurrentLineHardBreak();
                        StartNewLine();
                        pendingLeadingAdvance = 0;
                        pendingLeadingIsExpandable = true;
                        pendingLeadingIsTab = false;
                        pendingLeadingTabAlignment = PdfTabAlignment.Left;
                        pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                    } else if (nextWs != -1) {
                        bool hadTab = text[nextWs] == '\t';
                        pendingLeadingAdvance = hadTab ? CalculateTabAdvance(lineWidth, 0D, spaceW, tabAlignment, tabStopWidth) : spaceW;
                        pendingLeadingIsExpandable = !hadTab;
                        pendingLeadingIsTab = hadTab;
                        pendingLeadingTabAlignment = hadTab ? tabAlignment : PdfTabAlignment.Left;
                        pendingLeadingTabLeader = hadTab ? tabLeader : PdfTabLeaderStyle.None;
                    }
                    continue;
                }
                if (token.Length > 0 && pendingLeadingIsTab) {
                    pendingLeadingAdvance = CalculateTabAdvance(lineWidth, tokenW, spaceW, pendingLeadingTabAlignment, tabStopWidth, token, fontForRun, runFontSize, baseline);
                }
                needed = lastLine.Count == 0
                    ? (pendingLeadingIsTab ? pendingLeadingAdvance + tokenW : tokenW)
                    : pendingLeadingAdvance + tokenW;
                if (lineWidth + needed > currentMaxWidth && lastLine.Count > 0) {
                    StartNewLine();
                }
                if (token.Length > 0) {
                    bool needsLeadingSpace = pendingLeadingAdvance > 0 && (lineWidth > 0 || pendingLeadingIsTab);
                    double leadingAdvance = needsLeadingSpace ? pendingLeadingAdvance : 0;
                    double segmentWidth = tokenW + leadingAdvance;
                    var segmentLeader = needsLeadingSpace ? pendingLeadingTabLeader : PdfTabLeaderStyle.None;
                    lines[lines.Count - 1].Add(new RichSeg(token, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline, needsLeadingSpace, leadingAdvance, pendingLeadingIsExpandable, segmentLeader));
                    RegisterLineHeight(runFontSize);
                    lineWidth += segmentWidth;
                    pendingLeadingAdvance = 0;
                    pendingLeadingIsExpandable = true;
                    pendingLeadingIsTab = false;
                    pendingLeadingTabAlignment = PdfTabAlignment.Left;
                    pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                }
                if (hadNewline) {
                    MarkCurrentLineHardBreak();
                    StartNewLine();
                    pendingLeadingAdvance = 0;
                    pendingLeadingIsExpandable = true;
                    pendingLeadingIsTab = false;
                    pendingLeadingTabAlignment = PdfTabAlignment.Left;
                    pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                } else if (nextWs != -1) {
                    bool hadTab = text[nextWs] == '\t';
                    pendingLeadingAdvance = hadTab ? CalculateTabAdvance(lineWidth, 0D, spaceW, tabAlignment, tabStopWidth) : spaceW;
                    pendingLeadingIsExpandable = !hadTab;
                    pendingLeadingIsTab = hadTab;
                    pendingLeadingTabAlignment = hadTab ? tabAlignment : PdfTabAlignment.Left;
                    pendingLeadingTabLeader = hadTab ? tabLeader : PdfTabLeaderStyle.None;
                }
            }
        }
        if (lines.Count > 0 && lines[lines.Count - 1].Count == 0) { lines.RemoveAt(lines.Count - 1); }
        if (heights.Count < lines.Count) heights.Add(currentLineHeight);
        return (lines, heights);
    }

    private static void WriteRichParagraph(StringBuilder sb, RichParagraphBlock block, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, PdfOptions opts, double startY, double fontSize, double defaultLeading, System.Collections.Generic.List<LinkAnnotation> annots, double? xOverride = null, double? widthOverride = null, double? firstLineXOverride = null, double? firstLineWidthOverride = null) {
        double widthContent = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double widthUsed = widthOverride ?? widthContent;
        var underlines = new System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color)>();
        var strikes = new System.Collections.Generic.List<(double X1, double X2, double Y, PdfColor Color)>();
        var backgrounds = new System.Collections.Generic.List<(double X, double Y, double Width, double Height, PdfColor Color)>();

        void AddBackground(double x, double y, double width, double height, PdfColor color) {
            if (width <= 0.001D || height <= 0.001D) {
                return;
            }

            if (backgrounds.Count > 0) {
                var previous = backgrounds[backgrounds.Count - 1];
                if (previous.Color.Equals(color) &&
                    Math.Abs(previous.Y - y) <= 0.001D &&
                    Math.Abs(previous.Height - height) <= 0.001D &&
                    x <= previous.X + previous.Width + 0.25D) {
                    double left = Math.Min(previous.X, x);
                    double right = Math.Max(previous.X + previous.Width, x + width);
                    backgrounds[backgrounds.Count - 1] = (left, previous.Y, right - left, previous.Height, previous.Color);
                    return;
                }
            }

            backgrounds.Add((x, y, width, height, color));
        }

        double backgroundYOffset = 0D;
        double xOrigin = xOverride ?? opts.MarginLeft;
        for (int li = 0; li < lines.Count; li++) {
            double lineY = startY - backgroundYOffset;
            double lineWidthUsed = li == 0 ? firstLineWidthOverride ?? widthUsed : widthUsed;
            double lineXOrigin = li == 0 ? firstLineXOverride ?? xOrigin : xOrigin;
            var segs = lines[li];
            double baseLineW = 0;
            int gapsCount = 0;
            foreach (var seg in segs) {
                double w = MeasureRichText(seg.Text, seg.Font, seg.FontSize, seg.Baseline);
                if (seg.LeadingSpace) {
                    w += seg.LeadingAdvance > 0 ? seg.LeadingAdvance : MeasureRichText(" ", seg.Font, seg.FontSize, seg.Baseline);
                    if (seg.LeadingSpaceIsExpandable) {
                        gapsCount++;
                    }
                }

                baseLineW += w;
            }

            bool lineEndsWithHardBreak = segs.Any(seg => seg.EndsWithHardBreak);
            bool justify = block.Align == PdfAlign.Justify && !lineEndsWithHardBreak && li != lines.Count - 1 && gapsCount > 0 && lineWidthUsed > baseLineW;
            double wordSpacing = justify ? (lineWidthUsed - baseLineW) / gapsCount : 0;
            double lineWForAlign = justify ? lineWidthUsed : baseLineW;
            double dx = 0;
            if (block.Align == PdfAlign.Center) dx = Math.Max(0, (lineWidthUsed - lineWForAlign) / 2);
            else if (block.Align == PdfAlign.Right) dx = Math.Max(0, lineWidthUsed - lineWForAlign);

            double xCursor = dx;
            foreach (var s in segs) {
                double leadingAdvance = 0D;
                if (s.LeadingSpace) {
                    double baseGap = s.LeadingAdvance > 0 ? s.LeadingAdvance : MeasureRichText(" ", s.Font, s.FontSize, s.Baseline);
                    leadingAdvance = baseGap + (s.LeadingSpaceIsExpandable ? wordSpacing : 0);
                    xCursor += leadingAdvance;
                }

                double wSeg = MeasureRichText(s.Text, s.Font, s.FontSize, s.Baseline);
                if (s.BackgroundColor.HasValue && wSeg > 0) {
                    double runFontSize = EffectiveRichFontSize(s.FontSize, s.Baseline);
                    double textRise = TextRiseForBaseline(s.FontSize, s.Baseline);
                    double asc = GetAscender(s.Font, runFontSize);
                    double desc = GetDescender(s.Font, runFontSize);
                    double padX = Math.Max(1.4D, runFontSize * 0.14D);
                    double padY = Math.Max(0.45D, runFontSize * 0.05D);
                    double baselineY = lineY + textRise;
                    AddBackground(
                        lineXOrigin + xCursor - leadingAdvance - padX,
                        baselineY - desc - padY,
                        wSeg + leadingAdvance + (padX * 2D),
                        asc + desc + (padY * 2D),
                        s.BackgroundColor.Value);
                }

                xCursor += wSeg;
            }

            backgroundYOffset += li < lineHeights.Count ? lineHeights[li] : defaultLeading;
        }

        foreach (var bg in backgrounds) {
            new ContentStreamBuilder(sb)
                .SaveState()
                .FillColor(bg.Color)
                .Rectangle(bg.X, bg.Y, bg.Width, bg.Height)
                .FillPath()
                .RestoreState();
        }

        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .TextLeading(defaultLeading);

        double yOffset = 0D;
        for (int li = 0; li < lines.Count; li++) {
            double lineY = startY - yOffset;
            double lineWidthUsed = li == 0 ? firstLineWidthOverride ?? widthUsed : widthUsed;
            double lineXOrigin = li == 0 ? firstLineXOverride ?? xOrigin : xOrigin;
            var segs = lines[li];
            int segCount = segs.Count;
            double[] segWidths = segCount > 0 ? new double[segCount] : System.Array.Empty<double>();
            double baseLineW = 0;
            int gapsCount = 0;
            for (int si = 0; si < segCount; si++) {
                var seg = segs[si];
                double w = MeasureRichText(seg.Text, seg.Font, seg.FontSize, seg.Baseline);
                if (seg.LeadingSpace) {
                    w += seg.LeadingAdvance > 0 ? seg.LeadingAdvance : MeasureRichText(" ", seg.Font, seg.FontSize, seg.Baseline);
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
                .TextMatrix(lineXOrigin + dx, lineY)
                .WordSpacing(wordSpacing);

            double xCursor = dx;
            double currentTextRise = 0;
            for (int si = 0; si < segs.Count; si++) {
                var s = segs[si];
                string fontRes = GetStandardFontResourceName(s.Font, ChooseNormal(opts.DefaultFont));
                double runFontSize = EffectiveRichFontSize(s.FontSize, s.Baseline);
                double textRise = TextRiseForBaseline(s.FontSize, s.Baseline);
                content.Font(fontRes, runFontSize);
                if (Math.Abs(textRise - currentTextRise) > 0.0001) {
                    content.TextRise(textRise);
                    currentTextRise = textRise;
                }

                var color = s.Color ?? block.DefaultColor ?? opts.DefaultTextColor;
                content.FillColor(color ?? PdfColor.Black);
                if (s.LeadingSpace) {
                    double baseGap = s.LeadingAdvance > 0 ? s.LeadingAdvance : MeasureRichText(" ", s.Font, s.FontSize, s.Baseline);
                    double gap = baseGap + (s.LeadingSpaceIsExpandable ? wordSpacing : 0);

                    if (s.LeadingTabLeader != PdfTabLeaderStyle.None) {
                        string leader = BuildTabLeaderText(gap, s.Font, s.FontSize, s.Baseline, s.LeadingTabLeader);
                        if (leader.Length > 0) {
                            content
                                .TextMatrix(lineXOrigin + xCursor, lineY)
                                .ShowHexText(EncodeWinAnsiHex(leader));
                        }
                        xCursor += gap;
                        content.TextMatrix(lineXOrigin + xCursor, lineY);
                    } else if (!s.LeadingSpaceIsExpandable) {
                        content
                            .TextMatrix(lineXOrigin + xCursor, lineY)
                            .ShowHexText("20");
                        xCursor += gap;
                        content.TextMatrix(lineXOrigin + xCursor, lineY);
                    } else {
                        content.ShowHexText("20");
                        xCursor += gap;
                    }
                }
                double segmentStartX = xCursor;
                content.ShowHexText(EncodeWinAnsiHex(s.Text));
                double wSeg = MeasureRichText(s.Text, s.Font, s.FontSize, s.Baseline);
                double baselineY = lineY + textRise;

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
                    AddRichTextLinkAnnotation(annots, x1, y1, x2, y2, s.Uri, s.DestinationName, s.Contents);
                }
                xCursor += wSeg;
            }
            if (Math.Abs(currentTextRise) > 0.0001) {
                content.TextRise(0);
            }

            yOffset += li < lineHeights.Count ? lineHeights[li] : defaultLeading;
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

    private static void AddRichTextLinkAnnotation(System.Collections.Generic.List<LinkAnnotation> annots, double x1, double y1, double x2, double y2, string? uri, string? destinationName, string? contents) {
        if (annots.Count > 0) {
            LinkAnnotation previous = annots[annots.Count - 1];
            double gap = x1 - previous.X2;
            bool sameTarget =
                string.Equals(previous.Uri, uri, System.StringComparison.Ordinal) &&
                string.Equals(previous.DestinationName, destinationName, System.StringComparison.Ordinal) &&
                string.Equals(previous.Contents, contents, System.StringComparison.Ordinal);
            bool sameLine =
                Math.Abs(previous.Y1 - y1) <= 0.5D &&
                Math.Abs(previous.Y2 - y2) <= 0.5D;
            if (sameTarget && sameLine && gap >= -0.25D && gap <= 18D) {
                annots[annots.Count - 1] = new LinkAnnotation {
                    X1 = previous.X1,
                    Y1 = Math.Min(previous.Y1, y1),
                    X2 = Math.Max(previous.X2, x2),
                    Y2 = Math.Max(previous.Y2, y2),
                    Uri = uri,
                    DestinationName = destinationName,
                    Contents = contents
                };
                return;
            }
        }

        annots.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = uri, DestinationName = destinationName, Contents = contents });
    }

    private static void WriteClippedRichParagraph(StringBuilder sb, RichParagraphBlock block, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, PdfOptions opts, double startY, double fontSize, double defaultLeading, System.Collections.Generic.List<LinkAnnotation> annots, double clipX, double clipY, double clipWidth, double clipHeight, double? xOverride = null, double? widthOverride = null, double? firstLineXOverride = null, double? firstLineWidthOverride = null) {
        new ContentStreamBuilder(sb)
            .SaveState()
            .Rectangle(clipX, clipY, clipWidth, clipHeight)
            .ClipPath()
            .EndPath();

        WriteRichParagraph(sb, block, lines, lineHeights, opts, startY, fontSize, defaultLeading, annots, xOverride, widthOverride, firstLineXOverride, firstLineWidthOverride);

        new ContentStreamBuilder(sb)
            .RestoreState();
    }

    private static string BuildTabLeaderText(double gap, PdfStandardFont font, double fontSize, PdfTextBaseline baseline, PdfTabLeaderStyle leaderStyle) {
        string leaderGlyph = leaderStyle switch {
            PdfTabLeaderStyle.Dots => ".",
            PdfTabLeaderStyle.Hyphens => "-",
            PdfTabLeaderStyle.Underscores => "_",
            _ => string.Empty
        };

        if (leaderGlyph.Length == 0) {
            return string.Empty;
        }

        double glyphWidth = MeasureRichText(leaderGlyph, font, fontSize, baseline);
        if (glyphWidth <= 0 || gap <= glyphWidth * 3D) {
            return string.Empty;
        }

        int count = Math.Max(3, (int)Math.Floor(gap / glyphWidth));
        return new string(leaderGlyph[0], count);
    }
}
