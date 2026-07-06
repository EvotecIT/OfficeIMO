using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private const double DefaultParagraphTabStopWidth = 36D;
    private static readonly char[] TokenSplitChars = new[] { ' ', '\n', '\t' };
    private static readonly char[] HardLineSplitChars = new[] { '\n' };
    private static readonly char[] SoftLineSplitChars = new[] { ' ', '\t' };
    private static readonly char[] DecimalTabAnchorChars = new[] { '.', ',' };
    private static readonly char[] LongTokenDelimiterBreakChars = new[] { '-', '.', '_', '/', '\\', ':', '|' };
    private static string EscapeText(string s) => PdfSyntaxEscaper.EscapeLiteralContent(s);

    private static string EncodeWinAnsiHex(string s) {
        var bytes = PdfWinAnsiEncoding.Encode(s);
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        return sb.ToString();
    }

    private static string EncodeTextHex(string text, PdfStandardFont font, PdfOptions? options) {
        PdfTextEncodingDiagnostic? diagnostic = GetFirstTextEncodingDiagnostic(text, font, options);
        if (diagnostic != null) {
            throw CreateTextEncodingException(diagnostic, nameof(text));
        }

        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(font, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, fontProgram.FontDataSnapshot, fontName: fontProgram.FontName)
                : Array.Empty<PdfTextShapingDiagnostic>();
            options.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram));
            string glyphHex = fontProgram.EncodeTextAsGlyphHex(text, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot, options.RecordProviderShapedTextRun);
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, fontProgram.FontName, isOpenTypeCff: false);
            return glyphHex;
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(font, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            IReadOnlyList<PdfTextShapingDiagnostic> shapingDiagnostics = options.HasDiagnosticsReport
                ? PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, cffFontProgram.FontDataSnapshot, fontName: cffFontProgram.FontName)
                : Array.Empty<PdfTextShapingDiagnostic>();
            options.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, cffFontProgram));
            string glyphHex = cffFontProgram.EncodeTextAsGlyphHex(text, options.TextShapingModeSnapshot, options.TextShapingProviderSnapshot, options.RecordProviderShapedTextRun);
            options.AddTextShapingDiagnostics(shapingDiagnostics, text, cffFontProgram.FontName, isOpenTypeCff: true);
            return glyphHex;
        }

        if (options?.HasDiagnosticsReport == true) {
            options.AddTextShapingDiagnostics(PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text), text, deferProviderCoverable: false);
        }

        options?.AddTextDiagnostics(PdfTextDiagnostics.AnalyzeWinAnsiText(text));
        return EncodeWinAnsiHex(text);
    }

    private static PdfTextEncodingDiagnostic? GetFirstTextEncodingDiagnostic(string text, PdfStandardFont font, PdfOptions? options) {
        System.Collections.Generic.IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = options == null
            ? PdfTextDiagnostics.AnalyzeWinAnsiText(text, "generated text")
            : PdfTextDiagnostics.AnalyzeGeneratedText(text, options, font, "generated text");

        return diagnostics.Count == 0 ? null : diagnostics[0];
    }

    private static ArgumentException CreateTextEncodingException(PdfTextEncodingDiagnostic diagnostic, string paramName) {
        var exception = new ArgumentException(diagnostic.Message, paramName);
        exception.Data["code"] = diagnostic.Code;
        exception.Data["source"] = diagnostic.Source;
        exception.Data["index"] = diagnostic.Index;
        exception.Data["codePoint"] = diagnostic.CodePoint;
        exception.Data["text"] = diagnostic.Text;
        exception.Data["isControlCharacter"] = diagnostic.IsControlCharacter;
        if (!string.IsNullOrWhiteSpace(diagnostic.Location)) {
            exception.Data["location"] = diagnostic.Location;
        }

        if (!string.IsNullOrWhiteSpace(diagnostic.Encoding)) {
            exception.Data["encoding"] = diagnostic.Encoding;
        }

        if (!string.IsNullOrWhiteSpace(diagnostic.Remediation)) {
            exception.Data["remediation"] = diagnostic.Remediation;
        }

        return exception;
    }

    private static int GetScalarUtf16Length(string text, int index) {
        if (index < 0 || index >= text.Length) {
            throw new ArgumentOutOfRangeException(nameof(index), "Text scalar index must be inside the string.");
        }

        return char.IsHighSurrogate(text[index]) &&
            index + 1 < text.Length &&
            char.IsLowSurrogate(text[index + 1])
                ? 2
                : 1;
    }

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch) && index < text.Length && char.IsLowSurrogate(text[index])) {
            return char.ConvertToUtf32(ch, text[index++]);
        }

        return ch;
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

    private static System.Collections.Generic.List<string> WrapSimpleText(string text, double widthPts, PdfStandardFont font, double fontSize) =>
        WrapSimpleTextForOptions(text, widthPts, font, fontSize, options: null);

    private static System.Collections.Generic.List<string> WrapSimpleTextForOptions(string text, double widthPts, PdfStandardFont font, double fontSize, PdfOptions? options) {
        var hardLines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split(HardLineSplitChars, StringSplitOptions.None);
        var lines = new System.Collections.Generic.List<string>();
        double maxWidth = Math.Max(1D, widthPts);
        double spaceWidth = EstimateSimpleTextWidthForOptions(" ", font, fontSize, options);

        void FlushLine(StringBuilder current, ref double currentWidth) {
            if (current.Length > 0) {
                lines.Add(current.ToString());
                current.Clear();
                currentWidth = 0D;
            }
        }

        void AppendLongToken(string token, StringBuilder current, ref double currentWidth) {
            FlushLine(current, ref currentWidth);
            var softLineBreakChunks = TryBuildSoftLineBreakTokenChunks(
                token,
                options,
                part => EstimateSimpleTextWidthForOptions(part, font, fontSize, options),
                maxWidth,
                maxWidth);
            if (softLineBreakChunks != null) {
                AppendTokenChunks(softLineBreakChunks, current, ref currentWidth);
                return;
            }

            var multilingualChunks = TryBuildMultilingualTokenChunks(
                token,
                part => EstimateSimpleTextWidthForOptions(part, font, fontSize, options),
                maxWidth,
                maxWidth);
            if (multilingualChunks != null) {
                AppendTokenChunks(multilingualChunks, current, ref currentWidth);
                return;
            }

            for (int i = 0; i < token.Length; i++) {
                int scalarLength = GetScalarUtf16Length(token, i);
                string scalar = token.Substring(i, scalarLength);
                double characterWidth = EstimateSimpleTextWidthForOptions(scalar, font, fontSize, options);
                if (current.Length > 0 && currentWidth + characterWidth > maxWidth) {
                    FlushLine(current, ref currentWidth);
                }

                current.Append(scalar);
                currentWidth += characterWidth;
                i += scalarLength - 1;
            }
        }

        void AppendTokenChunks(System.Collections.Generic.IReadOnlyList<PdfTextTokenChunk> chunks, StringBuilder current, ref double currentWidth) {
            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                PdfTextTokenChunk chunk = chunks[chunkIndex];
                current.Append(chunk.Text);
                currentWidth += chunk.Width;
                if (chunkIndex + 1 < chunks.Count) {
                    FlushLine(current, ref currentWidth);
                }
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
                    double tokenWidth = EstimateSimpleTextWidthForOptions(token, font, fontSize, options);
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
            bool endsWithHardBreak = false,
            bool endsWithTextSeparator = false) {
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
            EndsWithTextSeparator = endsWithTextSeparator;
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

        public bool EndsWithTextSeparator { get; }

        public RichSeg WithEndsWithHardBreak() =>
            new RichSeg(Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, Uri, DestinationName, Contents, Font, FontSize, Baseline, LeadingSpace, LeadingAdvance, LeadingSpaceIsExpandable, LeadingTabLeader, true, true);

        public RichSeg WithEndsWithTextSeparator() =>
            new RichSeg(Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, Uri, DestinationName, Contents, Font, FontSize, Baseline, LeadingSpace, LeadingAdvance, LeadingSpaceIsExpandable, LeadingTabLeader, EndsWithHardBreak, true);

        public RichSeg WithoutLink() =>
            new RichSeg(Text, Bold, Italic, Underline, Strike, Color, BackgroundColor, null, null, null, Font, FontSize, Baseline, LeadingSpace, LeadingAdvance, LeadingSpaceIsExpandable, LeadingTabLeader, EndsWithHardBreak, EndsWithTextSeparator);
    }

    private static void MarkRichLineTextSeparator(System.Collections.Generic.IList<RichSeg> line) {
        if (line.Count == 0) {
            return;
        }

        int lastIndex = line.Count - 1;
        line[lastIndex] = line[lastIndex].WithEndsWithTextSeparator();
    }

    private static double MeasureRichText(string text, PdfStandardFont font, double fontSize, PdfOptions? options = null) =>
        EstimateSimpleTextWidthForOptions(text, font, fontSize, options);

    private static double EffectiveRichFontSize(double fontSize, PdfTextBaseline baseline) =>
        baseline == PdfTextBaseline.Normal ? fontSize : fontSize * 0.65;

    private static double TextRiseForBaseline(double fontSize, PdfTextBaseline baseline) => baseline switch {
        PdfTextBaseline.Superscript => fontSize * 0.35,
        PdfTextBaseline.Subscript => -fontSize * 0.18,
        _ => 0
    };

    private static double MeasureRichText(string text, PdfStandardFont font, double fontSize, PdfTextBaseline baseline, PdfOptions? options = null) =>
        EstimateSimpleTextWidthForOptions(text, font, EffectiveRichFontSize(fontSize, baseline), options);

    private static double MeasureRichLineWidth(System.Collections.Generic.IReadOnlyList<RichSeg> line, PdfOptions? options = null) {
        double width = 0D;
        for (int index = 0; index < line.Count; index++) {
            RichSeg segment = line[index];
            if (segment.LeadingSpace) {
                width += segment.LeadingAdvance > 0
                    ? segment.LeadingAdvance
                    : MeasureRichText(" ", segment.Font, segment.FontSize, segment.Baseline, options);
            }

            width += MeasureRichText(segment.Text, segment.Font, segment.FontSize, segment.Baseline, options);
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

    private static double CalculateTabAdvance(double lineWidth, double followingTextWidth, double spaceWidth, PdfTabAlignment alignment, double tabStopWidth = DefaultParagraphTabStopWidth, string followingText = "", PdfStandardFont followingFont = PdfStandardFont.Helvetica, double fontSize = 12D, PdfTextBaseline baseline = PdfTextBaseline.Normal, PdfOptions? options = null, double? maxWidth = null, PdfTabStop? explicitTabStop = null, double lineOriginOffset = 0D) {
        if (explicitTabStop == null && alignment == PdfTabAlignment.Left) {
            return CalculateDefaultTabAdvance(lineWidth, spaceWidth, tabStopWidth);
        }

        if (lineWidth < 0 || double.IsNaN(lineWidth) || double.IsInfinity(lineWidth) ||
            followingTextWidth < 0 || double.IsNaN(followingTextWidth) || double.IsInfinity(followingTextWidth) ||
            double.IsNaN(lineOriginOffset) || double.IsInfinity(lineOriginOffset)) {
            return spaceWidth;
        }

        if (explicitTabStop == null &&
            (tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth))) {
            return spaceWidth;
        }

        double? boundedMaxWidth = maxWidth.HasValue &&
            maxWidth.Value > 0 &&
            !double.IsNaN(maxWidth.Value) &&
            !double.IsInfinity(maxWidth.Value)
                ? maxWidth.Value
                : null;
        if (explicitTabStop != null) {
            alignment = explicitTabStop.Alignment;
        }

        double anchorWidth = alignment switch {
            PdfTabAlignment.Center => followingTextWidth / 2D,
            PdfTabAlignment.Right => followingTextWidth,
            PdfTabAlignment.DecimalSeparator => MeasureDecimalAnchorWidth(followingText, followingFont, fontSize, baseline, options),
            _ => 0D
        };
        double nextStop = explicitTabStop?.Position - lineOriginOffset ?? (Math.Floor(lineWidth / tabStopWidth) + 1D) * tabStopWidth;
        if (boundedMaxWidth.HasValue) {
            nextStop = Math.Min(nextStop, boundedMaxWidth.Value);
        }

        double advance = nextStop - anchorWidth - lineWidth;
        if (explicitTabStop != null) {
            return Math.Max(0D, advance);
        }

        if (advance < spaceWidth) {
            if (boundedMaxWidth.HasValue && nextStop >= boundedMaxWidth.Value) {
                return Math.Max(0D, advance);
            }

            double stopsToAdd = Math.Ceiling((spaceWidth - advance) / tabStopWidth);
            nextStop += Math.Max(1D, stopsToAdd) * tabStopWidth;
            if (boundedMaxWidth.HasValue) {
                nextStop = Math.Min(nextStop, boundedMaxWidth.Value);
            }

            advance = nextStop - anchorWidth - lineWidth;
            if (boundedMaxWidth.HasValue && nextStop >= boundedMaxWidth.Value) {
                return Math.Max(0D, advance);
            }
        }

        return Math.Max(spaceWidth, advance);
    }

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRuns(System.Collections.Generic.IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts = null, double tabStopWidth = DefaultParagraphTabStopWidth) =>
        WrapRichRunsCore(runs, maxWidthPts, fontSize, baseFont, lineHeight, firstLineWidthPts, tabStopWidth, options: null);

    private static PdfTabStop[]? NormalizeExplicitTabStops(System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops) {
        if (tabStops == null || tabStops.Count == 0) {
            return null;
        }

        return tabStops
            .Where(tabStop => tabStop.Position > 0 && !double.IsNaN(tabStop.Position) && !double.IsInfinity(tabStop.Position))
            .OrderBy(tabStop => tabStop.Position)
            .Select(tabStop => tabStop.Clone())
            .ToArray();
    }

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRunsCore(System.Collections.Generic.IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts, double tabStopWidth, PdfOptions? options, System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops = null) {
        return WrapRichRunsCoreWithFirstLineOrigin(runs, maxWidthPts, fontSize, baseFont, lineHeight, firstLineWidthPts, null, tabStopWidth, options, tabStops);
    }

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRunsCoreWithFirstLineOrigin(System.Collections.Generic.IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts, double? firstLineOriginOffsetPts, double tabStopWidth, PdfOptions? options, System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops = null) {
        System.Collections.Generic.IEnumerable<TextRun> effectiveRuns = NormalizeFallbackRuns(runs, baseFont, options);
        PdfTabStop[]? explicitTabStops = NormalizeExplicitTabStops(tabStops);
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new() };
        var heights = new System.Collections.Generic.List<double>();
        double lineWidth = 0;
        double pendingLeadingAdvance = 0;
        bool pendingLeadingIsExpandable = true;
        bool pendingLeadingIsTab = false;
        PdfTabAlignment pendingLeadingTabAlignment = PdfTabAlignment.Left;
        PdfTabLeaderStyle pendingLeadingTabLeader = PdfTabLeaderStyle.None;
        PdfTabStop? pendingLeadingTabStop = null;
        int nextExplicitTabStopIndex = 0;
        double lineHeightRatio = fontSize > 0 ? lineHeight / fontSize : 1.2D;
        double currentLineHeight = lineHeight;
        double CurrentMaxWidth() => lines.Count == 1 ? firstLineWidthPts ?? maxWidthPts : maxWidthPts;
        double CurrentLineOriginOffset() => lines.Count == 1 ? firstLineOriginOffsetPts ?? 0D : 0D;
        void RegisterLineHeight(double runFontSize) {
            currentLineHeight = Math.Max(currentLineHeight, runFontSize * lineHeightRatio);
        }

        void StartNewLine() {
            heights.Add(currentLineHeight);
            lines.Add(new());
            lineWidth = 0;
            currentLineHeight = lineHeight;
            nextExplicitTabStopIndex = 0;
        }

        PdfTabStop? ResolveNextExplicitTabStop() {
            if (explicitTabStops == null || explicitTabStops.Length == 0) {
                return null;
            }

            while (nextExplicitTabStopIndex < explicitTabStops.Length &&
                   explicitTabStops[nextExplicitTabStopIndex].Position <= CurrentLineOriginOffset() + lineWidth + 0.001D) {
                nextExplicitTabStopIndex++;
            }

            if (nextExplicitTabStopIndex >= explicitTabStops.Length) {
                return null;
            }

            return explicitTabStops[nextExplicitTabStopIndex++];
        }

        void ResolvePendingLeadingTabForCurrentLine(double followingTextWidth, double spaceW, string followingText, PdfStandardFont followingFont, double followingFontSize, PdfTextBaseline followingBaseline) {
            PdfTabAlignment fallbackAlignment = pendingLeadingTabAlignment;
            PdfTabLeaderStyle fallbackLeader = pendingLeadingTabLeader;
            PdfTabStop? explicitTabStop = ResolveNextExplicitTabStop();
            pendingLeadingTabAlignment = explicitTabStop?.Alignment ?? fallbackAlignment;
            pendingLeadingTabLeader = explicitTabStop?.Leader ?? fallbackLeader;
            pendingLeadingTabStop = explicitTabStop;
            pendingLeadingAdvance = CalculateTabAdvance(lineWidth, followingTextWidth, spaceW, pendingLeadingTabAlignment, tabStopWidth, followingText, followingFont, followingFontSize, followingBaseline, options, CurrentMaxWidth(), pendingLeadingTabStop, CurrentLineOriginOffset());
        }

        void ResetPendingLeading() {
            pendingLeadingAdvance = 0;
            pendingLeadingIsExpandable = true;
            pendingLeadingIsTab = false;
            pendingLeadingTabAlignment = PdfTabAlignment.Left;
            pendingLeadingTabLeader = PdfTabLeaderStyle.None;
            pendingLeadingTabStop = null;
        }

        void SetPendingSeparator(bool hadTab, double spaceW, PdfTabAlignment tabAlignment, PdfTabLeaderStyle tabLeader) {
            if (!hadTab) {
                pendingLeadingAdvance = spaceW;
                pendingLeadingIsExpandable = true;
                pendingLeadingIsTab = false;
                pendingLeadingTabAlignment = PdfTabAlignment.Left;
                pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                pendingLeadingTabStop = null;
                return;
            }

            PdfTabStop? explicitTabStop = ResolveNextExplicitTabStop();
            pendingLeadingTabAlignment = explicitTabStop?.Alignment ?? tabAlignment;
            pendingLeadingTabLeader = explicitTabStop?.Leader ?? tabLeader;
            pendingLeadingTabStop = explicitTabStop;
            pendingLeadingAdvance = CalculateTabAdvance(lineWidth, 0D, spaceW, pendingLeadingTabAlignment, tabStopWidth, options: options, maxWidth: CurrentMaxWidth(), explicitTabStop: pendingLeadingTabStop, lineOriginOffset: CurrentLineOriginOffset());
            pendingLeadingIsExpandable = false;
            pendingLeadingIsTab = true;
        }

        void MarkCurrentLineHardBreak() {
            var currentLine = lines[lines.Count - 1];
            if (currentLine.Count == 0) {
                return;
            }

            var lastSegment = currentLine[currentLine.Count - 1];
            currentLine[currentLine.Count - 1] = lastSegment.WithEndsWithHardBreak();
        }

        foreach (var run in effectiveRuns) {
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
            double spaceW = MeasureRichText(" ", fontForRun, runFontSize, baseline, options);
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
                double tokenW = MeasureRichText(token, fontForRun, runFontSize, baseline, options);
                var lastLine = lines[lines.Count - 1];
                double needed = lastLine.Count == 0 ? tokenW : pendingLeadingAdvance + tokenW;
                double currentMaxWidth = CurrentMaxWidth();

                if (tokenW > currentMaxWidth) {
                    if (lastLine.Count > 0) { StartNewLine(); lastLine = lines[lines.Count - 1]; }
                    ResetPendingLeading();
                    if (TryAppendSoftLineBreakLongToken(token, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline)) {
                        if (hadNewline) {
                            MarkCurrentLineHardBreak();
                            StartNewLine();
                            ResetPendingLeading();
                        } else if (nextWs != -1) {
                            bool hadTab = text[nextWs] == '\t';
                            SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                        }
                        continue;
                    }

                    if (TryAppendDelimitedLongToken(token, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline)) {
                        if (hadNewline) {
                            MarkCurrentLineHardBreak();
                            StartNewLine();
                            ResetPendingLeading();
                        } else if (nextWs != -1) {
                            bool hadTab = text[nextWs] == '\t';
                            SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                        }
                        continue;
                    }

                    if (TryAppendHyphenatedLongToken(token, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline)) {
                        if (hadNewline) {
                            MarkCurrentLineHardBreak();
                            StartNewLine();
                            ResetPendingLeading();
                        } else if (nextWs != -1) {
                            bool hadTab = text[nextWs] == '\t';
                            SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                        }
                        continue;
                    }

                    if (TryAppendMultilingualLongToken(token, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline)) {
                        if (hadNewline) {
                            MarkCurrentLineHardBreak();
                            StartNewLine();
                            ResetPendingLeading();
                        } else if (nextWs != -1) {
                            bool hadTab = text[nextWs] == '\t';
                            SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                        }
                        continue;
                    }

                    int pos = 0;
                    while (pos < token.Length) {
                        int take = 0;
                        double chunkW = 0;
                        currentMaxWidth = CurrentMaxWidth();
                        while (pos + take < token.Length) {
                            int scalarLength = GetScalarUtf16Length(token, pos + take);
                            string scalar = token.Substring(pos + take, scalarLength);
                            double charW = MeasureRichText(scalar, fontForRun, runFontSize, baseline, options);
                            if (take > 0 && chunkW + charW > currentMaxWidth) {
                                break;
                            }

                            chunkW += charW;
                            take += scalarLength;
                            if (chunkW >= currentMaxWidth) {
                                break;
                            }
                        }

                        if (take == 0) {
                            take = GetScalarUtf16Length(token, pos);
                            chunkW = MeasureRichText(token.Substring(pos, take), fontForRun, runFontSize, baseline, options);
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
                        ResetPendingLeading();
                    } else if (nextWs != -1) {
                        bool hadTab = text[nextWs] == '\t';
                        SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                    }
                    continue;
                }
                if (token.Length > 0 && pendingLeadingIsTab) {
                    pendingLeadingAdvance = CalculateTabAdvance(lineWidth, tokenW, spaceW, pendingLeadingTabAlignment, tabStopWidth, token, fontForRun, runFontSize, baseline, options, CurrentMaxWidth(), pendingLeadingTabStop, CurrentLineOriginOffset());
                }
                needed = lastLine.Count == 0
                    ? (pendingLeadingIsTab ? pendingLeadingAdvance + tokenW : tokenW)
                    : pendingLeadingAdvance + tokenW;
                if (lineWidth + needed > currentMaxWidth && lastLine.Count > 0) {
                    if (pendingLeadingAdvance > 0D) {
                        MarkRichLineTextSeparator(lastLine);
                    }

                    StartNewLine();
                    if (token.Length > 0 && pendingLeadingIsTab) {
                        ResolvePendingLeadingTabForCurrentLine(tokenW, spaceW, token, fontForRun, runFontSize, baseline);
                    }
                }
                if (token.Length > 0) {
                    bool needsLeadingSpace = pendingLeadingAdvance > 0 && (lineWidth > 0 || pendingLeadingIsTab);
                    double leadingAdvance = needsLeadingSpace ? pendingLeadingAdvance : 0;
                    double segmentWidth = tokenW + leadingAdvance;
                    var segmentLeader = needsLeadingSpace ? pendingLeadingTabLeader : PdfTabLeaderStyle.None;
                    lines[lines.Count - 1].Add(new RichSeg(token, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline, needsLeadingSpace, leadingAdvance, pendingLeadingIsExpandable, segmentLeader));
                    RegisterLineHeight(runFontSize);
                    lineWidth += segmentWidth;
                    ResetPendingLeading();
                }
                if (hadNewline) {
                    MarkCurrentLineHardBreak();
                    StartNewLine();
                    ResetPendingLeading();
                } else if (nextWs != -1) {
                    bool hadTab = text[nextWs] == '\t';
                    SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                }
            }
        }
        if (lines.Count > 0 && lines[lines.Count - 1].Count == 0) { lines.RemoveAt(lines.Count - 1); }
        if (heights.Count < lines.Count) heights.Add(currentLineHeight);
        return (lines, heights);

        bool TryAppendSoftLineBreakLongToken(
            string token,
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
            double runFontSize,
            PdfTextBaseline baseline) {
            var chunks = TryBuildSoftLineBreakTokenChunks(
                token,
                options,
                part => MeasureRichText(part, font, runFontSize, baseline, options),
                CurrentMaxWidth(),
                maxWidthPts);

            if (chunks == null) {
                return false;
            }

            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                PdfTextTokenChunk chunk = chunks[chunkIndex];
                lines[lines.Count - 1].Add(new RichSeg(chunk.Text, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline));
                RegisterLineHeight(runFontSize);
                lineWidth += chunk.Width;
                if (chunkIndex + 1 < chunks.Count) {
                    StartNewLine();
                }
            }

            return true;
        }

        bool TryAppendHyphenatedLongToken(
            string token,
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
            double runFontSize,
            PdfTextBaseline baseline) {
            int[] breakpoints = GetValidHyphenationBreakpoints(token, options);
            if (breakpoints.Length == 0) {
                return false;
            }

            int position = 0;
            var plannedChunks = new System.Collections.Generic.List<(string Text, double Width)>();
            while (position < token.Length) {
                int selectedBreak = -1;
                string selectedText = string.Empty;
                double selectedWidth = 0D;
                double maxWidthForChunk = plannedChunks.Count == 0 ? CurrentMaxWidth() : maxWidthPts;
                int[] candidates = breakpoints
                    .Where(point => point > position)
                    .Concat(new[] { token.Length })
                    .Distinct()
                    .OrderBy(point => point)
                    .ToArray();

                foreach (int candidate in candidates) {
                    bool finalChunk = candidate >= token.Length;
                    string chunkText = token.Substring(position, candidate - position);
                    if (!finalChunk) {
                        chunkText += "-";
                    }

                    if (chunkText.Length == 0) {
                        continue;
                    }

                    double chunkWidth = MeasureRichText(chunkText, font, runFontSize, baseline, options);
                    if (chunkWidth <= maxWidthForChunk || selectedBreak < 0) {
                        if (chunkWidth <= maxWidthForChunk) {
                            selectedBreak = candidate;
                            selectedText = chunkText;
                            selectedWidth = chunkWidth;
                        }
                    }

                    if (chunkWidth > maxWidthForChunk && selectedBreak >= 0) {
                        break;
                    }
                }

                if (selectedBreak <= position || selectedText.Length == 0) {
                    return false;
                }

                plannedChunks.Add((selectedText, selectedWidth));
                position = selectedBreak;
            }

            for (int chunkIndex = 0; chunkIndex < plannedChunks.Count; chunkIndex++) {
                (string selectedText, double selectedWidth) = plannedChunks[chunkIndex];
                lines[lines.Count - 1].Add(new RichSeg(selectedText, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline));
                RegisterLineHeight(runFontSize);
                lineWidth += selectedWidth;
                if (chunkIndex < plannedChunks.Count - 1) {
                    StartNewLine();
                }
            }

            return true;
        }

        bool TryAppendDelimitedLongToken(
            string token,
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
            double runFontSize,
            PdfTextBaseline baseline) {
            int[] breakpoints = GetValidLongTokenDelimiterBreakpoints(token);
            if (breakpoints.Length == 0) {
                return false;
            }

            int position = 0;
            var plannedChunks = new System.Collections.Generic.List<(string Text, double Width)>();
            while (position < token.Length) {
                int selectedBreak = -1;
                string selectedText = string.Empty;
                double selectedWidth = 0D;
                double maxWidthForChunk = plannedChunks.Count == 0 ? CurrentMaxWidth() : maxWidthPts;
                if (TryPlanDelimiterBoundedWordGroup(maxWidthForChunk, out selectedText, out selectedWidth)) {
                    selectedBreak = position + selectedText.Length;
                    plannedChunks.Add((selectedText, selectedWidth));
                    position = selectedBreak;
                    continue;
                }

                int[] candidates = breakpoints
                    .Where(point => point > position)
                    .Concat(new[] { token.Length })
                    .Distinct()
                    .OrderBy(point => point)
                    .ToArray();

                foreach (int candidate in candidates) {
                    string chunkText = token.Substring(position, candidate - position);
                    if (chunkText.Length == 0) {
                        continue;
                    }

                    double chunkWidth = MeasureRichText(chunkText, font, runFontSize, baseline, options);
                    if (chunkWidth <= maxWidthForChunk || selectedBreak < 0) {
                        if (chunkWidth <= maxWidthForChunk) {
                            selectedBreak = candidate;
                            selectedText = chunkText;
                            selectedWidth = chunkWidth;
                        }
                    }

                    if (chunkWidth > maxWidthForChunk && selectedBreak >= 0) {
                        break;
                    }
                }

                if (selectedBreak > position &&
                    selectedBreak < token.Length &&
                    CanExtendDelimitedIdentifierChunk(token, selectedBreak)) {
                    TryExtendIdentifierChunkToAvailableWidth(maxWidthForChunk, ref selectedText, ref selectedWidth, ref selectedBreak);
                }

                if (selectedBreak <= position || selectedText.Length == 0) {
                    if (!TryPlanCharacterChunkToNextDelimiter(maxWidthForChunk, out selectedText, out selectedWidth)) {
                        return false;
                    }

                    selectedBreak = position + selectedText.Length;
                }

                plannedChunks.Add((selectedText, selectedWidth));
                position = selectedBreak;
            }

            for (int chunkIndex = 0; chunkIndex < plannedChunks.Count; chunkIndex++) {
                (string selectedText, double selectedWidth) = plannedChunks[chunkIndex];
                lines[lines.Count - 1].Add(new RichSeg(selectedText, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline));
                RegisterLineHeight(runFontSize);
                lineWidth += selectedWidth;
                if (chunkIndex < plannedChunks.Count - 1) {
                    StartNewLine();
                }
            }

            return true;

            bool TryPlanDelimiterBoundedWordGroup(double maxWidthForChunk, out string selectedText, out double selectedWidth) {
                selectedText = string.Empty;
                selectedWidth = 0D;
                if (position >= token.Length || !IsLongTokenDelimiterBreakChar(token[position])) {
                    return false;
                }

                int nextDelimiterIndex = -1;
                for (int index = position + 1; index < token.Length; index++) {
                    if (IsLongTokenDelimiterBreakChar(token[index])) {
                        nextDelimiterIndex = index;
                        break;
                    }
                }

                if (nextDelimiterIndex <= position + 1) {
                    return false;
                }

                string candidate = token.Substring(position, nextDelimiterIndex - position + 1);
                if (!IsDelimiterBoundedWordGroup(candidate)) {
                    return false;
                }

                double candidateWidth = MeasureRichText(candidate, font, runFontSize, baseline, options);
                if (candidateWidth <= maxWidthForChunk ||
                    CanKeepDelimitedWordSegmentOverSoftLimit(candidate, candidateWidth, maxWidthForChunk, allowBoundaryDelimiters: true)) {
                    selectedText = candidate;
                    selectedWidth = candidateWidth;
                    return true;
                }

                return false;
            }

            bool TryPlanCharacterChunkToNextDelimiter(double maxWidthForChunk, out string selectedText, out double selectedWidth) {
                selectedText = string.Empty;
                selectedWidth = 0D;
                int segmentEnd = breakpoints.FirstOrDefault(point => point > position);
                if (segmentEnd <= position) {
                    segmentEnd = token.Length;
                }

                if (segmentEnd < token.Length && segmentEnd > position + 1 && IsLongTokenDelimiterBreakChar(token[segmentEnd - 1])) {
                    string textWithoutTrailingDelimiter = token.Substring(position, segmentEnd - position - 1);
                    if (textWithoutTrailingDelimiter.Length > 0) {
                        double widthWithoutTrailingDelimiter = MeasureRichText(textWithoutTrailingDelimiter, font, runFontSize, baseline, options);
                        if (widthWithoutTrailingDelimiter <= maxWidthForChunk ||
                            CanKeepDelimitedWordSegmentOverSoftLimit(textWithoutTrailingDelimiter, widthWithoutTrailingDelimiter, maxWidthForChunk, allowBoundaryDelimiters: false)) {
                            selectedText = textWithoutTrailingDelimiter;
                            selectedWidth = widthWithoutTrailingDelimiter;
                            return true;
                        }
                    }
                }

                int take = 0;
                double chunkWidth = 0D;
                while (position + take < segmentEnd) {
                    int scalarLength = GetScalarUtf16Length(token, position + take);
                    string scalar = token.Substring(position + take, scalarLength);
                    double scalarWidth = MeasureRichText(scalar, font, runFontSize, baseline, options);
                    if (take > 0 && chunkWidth + scalarWidth > maxWidthForChunk) {
                        break;
                    }

                    chunkWidth += scalarWidth;
                    take += scalarLength;
                    if (chunkWidth >= maxWidthForChunk) {
                        break;
                    }
                }

                if (take == 0) {
                    return false;
                }

                selectedText = token.Substring(position, take);
                selectedWidth = chunkWidth;
                return true;
            }

            void TryExtendIdentifierChunkToAvailableWidth(double maxWidthForChunk, ref string selectedText, ref double selectedWidth, ref int selectedBreak) {
                int extendedBreak = selectedBreak;
                string extendedText = selectedText;
                double extendedWidth = selectedWidth;
                while (extendedBreak < token.Length && IsIdentifierContinuationChar(token[extendedBreak])) {
                    int scalarLength = GetScalarUtf16Length(token, extendedBreak);
                    string candidateText = token.Substring(position, extendedBreak + scalarLength - position);
                    double candidateWidth = MeasureRichText(candidateText, font, runFontSize, baseline, options);
                    if (candidateWidth > maxWidthForChunk) {
                        break;
                    }

                    extendedBreak += scalarLength;
                    extendedText = candidateText;
                    extendedWidth = candidateWidth;
                }

                if (extendedBreak > selectedBreak) {
                    selectedBreak = extendedBreak;
                    selectedText = extendedText;
                    selectedWidth = extendedWidth;
                }
            }

            bool CanKeepDelimitedWordSegmentOverSoftLimit(string text, double width, double maxWidth, bool allowBoundaryDelimiters) {
                if (maxWidth <= 0D || width <= maxWidth) {
                    return false;
                }

                bool validSegment = allowBoundaryDelimiters
                    ? IsDelimiterBoundedWordGroup(text)
                    : IsDelimitedWordSegment(text);
                if (!validSegment) {
                    return false;
                }

                double widestScalar = 0D;
                for (int offset = 0; offset < text.Length;) {
                    int scalarLength = GetScalarUtf16Length(text, offset);
                    string scalar = text.Substring(offset, scalarLength);
                    widestScalar = Math.Max(widestScalar, MeasureRichText(scalar, font, runFontSize, baseline, options));
                    offset += scalarLength;
                }

                double overflow = width - maxWidth;
                return widestScalar > 0D && overflow <= widestScalar + 0.25D;
            }

            static bool IsDelimitedWordSegment(string text) {
                if (text.Length < 5) {
                    return false;
                }

                for (int index = 0; index < text.Length; index++) {
                    if (!char.IsLetter(text[index])) {
                        return false;
                    }
                }

                return true;
            }

            static bool IsDelimiterBoundedWordGroup(string text) {
                if (text.Length < 4 ||
                    !IsLongTokenDelimiterBreakChar(text[0]) ||
                    !IsLongTokenDelimiterBreakChar(text[text.Length - 1])) {
                    return false;
                }

                for (int index = 1; index < text.Length - 1; index++) {
                    if (!char.IsLetter(text[index])) {
                        return false;
                    }
                }

                return true;
            }

            static bool CanExtendDelimitedIdentifierChunk(string token, int breakIndex) {
                char delimiter = token[breakIndex - 1];
                if (breakIndex >= token.Length) {
                    return false;
                }

                return (delimiter == '_' || delimiter == '/' || delimiter == '\\') &&
                    IsIdentifierContinuationChar(token[breakIndex]);
            }

            static bool IsIdentifierContinuationChar(char value) => char.IsLetterOrDigit(value);
        }

        bool TryAppendMultilingualLongToken(
            string token,
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
            double runFontSize,
            PdfTextBaseline baseline) {
            var chunks = TryBuildMultilingualTokenChunks(
                token,
                part => MeasureRichText(part, font, runFontSize, baseline, options),
                CurrentMaxWidth(),
                maxWidthPts);

            if (chunks == null) {
                return false;
            }

            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                PdfTextTokenChunk chunk = chunks[chunkIndex];
                lines[lines.Count - 1].Add(new RichSeg(chunk.Text, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline));
                RegisterLineHeight(runFontSize);
                lineWidth += chunk.Width;
                if (chunkIndex + 1 < chunks.Count) {
                    StartNewLine();
                }
            }

            return true;
        }
    }

    private static int[] GetValidHyphenationBreakpoints(string token, PdfOptions? options) {
        PdfTextHyphenationCallback? callback = options?.TextHyphenationCallbackSnapshot;
        if (callback == null || string.IsNullOrEmpty(token)) {
            return Array.Empty<int>();
        }

        System.Collections.Generic.IReadOnlyList<int>? points = callback(token);
        if (points == null || points.Count == 0) {
            return Array.Empty<int>();
        }

        return points
            .Where(point => IsValidTokenBreakIndex(token, point))
            .Distinct()
            .OrderBy(point => point)
            .ToArray();
    }

    private static bool IsValidTokenBreakIndex(string token, int index) =>
        index > 0 &&
        index < token.Length &&
        !(index > 0 && index < token.Length && char.IsHighSurrogate(token[index - 1]) && char.IsLowSurrogate(token[index]));

    private static int[] GetValidLongTokenDelimiterBreakpoints(string token) {
        if (string.IsNullOrEmpty(token)) {
            return Array.Empty<int>();
        }

        return token
            .Select((ch, index) => IsLongTokenDelimiterBreakChar(ch) ? index + 1 : -1)
            .Where(point => IsValidTokenBreakIndex(token, point))
            .Distinct()
            .OrderBy(point => point)
            .ToArray();
    }

    private static bool IsLongTokenDelimiterBreakChar(char value) =>
        Array.IndexOf(LongTokenDelimiterBreakChars, value) >= 0;

    private static System.Collections.Generic.IReadOnlyList<TextRun> NormalizeFallbackRuns(System.Collections.Generic.IEnumerable<TextRun> runs, PdfStandardFont baseFont, PdfOptions? options) {
        Guard.NotNull(runs, nameof(runs));
        PdfEmbeddedFontFallbackSet? fallbackSet = options?.EmbeddedFontFallbacksSnapshot;
        if (fallbackSet == null) {
            return runs as System.Collections.Generic.IReadOnlyList<TextRun> ?? runs.ToArray();
        }

        var normalized = new System.Collections.Generic.List<TextRun>();
        foreach (TextRun run in runs) {
            if (CanWriteRunWithSelectedFont(run, baseFont, options)) {
                normalized.Add(run);
                continue;
            }

            PdfTextShapingMode shapingMode = options?.TextShapingModeSnapshot ?? PdfTextShapingMode.UnicodeScalar;
            if (fallbackSet.TryPlanTextRuns(run.Text, out System.Collections.Generic.IReadOnlyList<TextRun> plannedRuns, styleTemplate: run, shapingMode: shapingMode) ||
                TryPlanFallbackRunsPreservingSelectedFont(run, baseFont, options, fallbackSet, out plannedRuns)) {
                normalized.AddRange(plannedRuns);
            } else {
                normalized.Add(run);
            }
        }

        return normalized;
    }

    private static bool TryPlanFallbackRunsPreservingSelectedFont(
        TextRun run,
        PdfStandardFont baseFont,
        PdfOptions? options,
        PdfEmbeddedFontFallbackSet fallbackSet,
        out System.Collections.Generic.IReadOnlyList<TextRun> plannedRuns) {
        plannedRuns = Array.Empty<TextRun>();
        string text = run.Text ?? string.Empty;
        if (text.Length == 0 || IsLayoutControlRun(run)) {
            plannedRuns = new[] { run };
            return true;
        }

        PdfStandardFont fontForRun = ResolveFontForRun(run, baseFont);
        var runs = new System.Collections.Generic.List<TextRun>();
        int selectedStart = -1;

        void FlushSelected(int endIndex) {
            if (selectedStart < 0 || endIndex <= selectedStart) {
                return;
            }

            runs.Add(CreateStyledTextRun(text.Substring(selectedStart, endIndex - selectedStart), run, run.Font));
            selectedStart = -1;
        }

        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            int scalar = ReadScalar(text, ref index);
            if (scalar == '\n' || scalar == '\r' || scalar == '\t') {
                FlushSelected(scalarStart);
                if (scalar == '\t') {
                    runs.Add(TextRun.Tab(run.TabLeader, run.TabAlignment));
                } else {
                    runs.Add(TextRun.LineBreak());
                    if (scalar == '\r' && index < text.Length && text[index] == '\n') {
                        index++;
                    }
                }

                continue;
            }

            if (TryGetSelectedTextLength(text, scalarStart, fontForRun, options, out int selectedLength)) {
                if (selectedStart < 0) {
                    selectedStart = scalarStart;
                }

                index = scalarStart + selectedLength;
                continue;
            }

            FlushSelected(scalarStart);
            string scalarText = text.Substring(scalarStart, index - scalarStart);
            PdfTextShapingMode shapingMode = options?.TextShapingModeSnapshot ?? PdfTextShapingMode.UnicodeScalar;
            if (!fallbackSet.TryPlanTextRuns(scalarText, out System.Collections.Generic.IReadOnlyList<TextRun> fallbackRuns, styleTemplate: run, shapingMode: shapingMode)) {
                plannedRuns = Array.Empty<TextRun>();
                return false;
            }

            runs.AddRange(fallbackRuns);
        }

        FlushSelected(text.Length);
        plannedRuns = runs.AsReadOnly();
        return true;
    }

    private static TextRun CreateStyledTextRun(string text, TextRun styleTemplate, PdfStandardFont? font) {
        bool keepLink = !string.IsNullOrWhiteSpace(text) &&
            (styleTemplate.LinkUri != null || styleTemplate.LinkDestinationName != null);

        return new TextRun(
            text,
            styleTemplate.Bold,
            styleTemplate.Underline,
            styleTemplate.Color,
            styleTemplate.Italic,
            styleTemplate.Strike,
            styleTemplate.FontSize,
            font,
            keepLink ? styleTemplate.LinkUri : null,
            keepLink ? styleTemplate.LinkContents : null,
            styleTemplate.Baseline,
            keepLink ? styleTemplate.LinkDestinationName : null,
            backgroundColor: styleTemplate.BackgroundColor);
    }

    private static bool CanWriteRunWithSelectedFont(TextRun run, PdfStandardFont baseFont, PdfOptions? options) {
        string text = run.Text ?? string.Empty;
        if (text.Length == 0 || IsLayoutControlRun(run)) {
            return true;
        }

        PdfStandardFont fontForRun = ResolveFontForRun(run, baseFont);
        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(fontForRun, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            return CanWriteWithEmbeddedFont(text, fontProgram, options.TextShapingModeSnapshot);
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(fontForRun, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            return CanWriteWithEmbeddedFont(text, cffFontProgram, options.TextShapingModeSnapshot);
        }

        return PdfWinAnsiEncoding.CanEncode(text, out _);
    }

    private static PdfStandardFont ResolveFontForRun(TextRun run, PdfStandardFont baseFont) {
        PdfStandardFont runBaseFont = run.Font.HasValue ? ChooseNormal(run.Font.Value) : baseFont;
        return (run.Bold && run.Italic)
            ? ChooseBoldItalic(runBaseFont)
            : run.Bold
                ? ChooseBold(runBaseFont)
                : run.Italic
                    ? ChooseItalic(runBaseFont)
                    : runBaseFont;
    }

    private static bool TryGetSelectedTextLength(string text, int index, PdfStandardFont fontForRun, PdfOptions? options, out int length) {
        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(fontForRun, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            return TryGetCoveredTextLength(text, index, fontProgram, options.TextShapingModeSnapshot, out length);
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(fontForRun, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            return TryGetCoveredTextLength(text, index, cffFontProgram, options.TextShapingModeSnapshot, out length);
        }

        int endIndex = index;
        _ = ReadScalar(text, ref endIndex);
        length = endIndex - index;
        return PdfWinAnsiEncoding.CanEncode(text.Substring(index, length), out _);
    }

    private static bool CanWriteWithEmbeddedFont(string text, PdfTrueTypeFontProgram fontProgram, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) {
        int index = 0;
        while (index < text.Length) {
            int scalarStart = index;
            if (shapingMode == PdfTextShapingMode.LatinLigatures &&
                PdfLatinLigatureSubstitution.TryGetPresentationLigature(text, scalarStart, out int ligatureScalar, out int ligatureLength) &&
                fontProgram.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
                ligatureGlyphId > 0) {
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (scalar == '\n' || scalar == '\r' || scalar == '\t') {
                continue;
            }

            if (scalar < ' ' || scalar == '\u007F') {
                return false;
            }

            if (!fontProgram.TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetCoveredTextLength(string text, int index, PdfTrueTypeFontProgram fontProgram, PdfTextShapingMode shapingMode, out int length) {
        if (shapingMode == PdfTextShapingMode.LatinLigatures &&
            PdfLatinLigatureSubstitution.TryGetPresentationLigature(text, index, out int ligatureScalar, out length) &&
            fontProgram.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
            ligatureGlyphId > 0) {
            return true;
        }

        int endIndex = index;
        int scalar = ReadScalar(text, ref endIndex);
        length = endIndex - index;
        return fontProgram.TryGetGlyphId(scalar, out int glyphId) && glyphId > 0;
    }

    private static bool CanWriteWithEmbeddedFont(string text, PdfOpenTypeCffFontProgram fontProgram, PdfTextShapingMode shapingMode = PdfTextShapingMode.UnicodeScalar) {
        int index = 0;
        while (index < text.Length) {
            int scalarStart = index;
            if (shapingMode == PdfTextShapingMode.LatinLigatures &&
                PdfLatinLigatureSubstitution.TryGetPresentationLigature(text, scalarStart, out int ligatureScalar, out int ligatureLength) &&
                fontProgram.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
                ligatureGlyphId > 0) {
                index += ligatureLength;
                continue;
            }

            int scalar = ReadScalar(text, ref index);
            if (scalar == '\n' || scalar == '\r' || scalar == '\t') {
                continue;
            }

            if (scalar < ' ' || scalar == '\u007F') {
                return false;
            }

            if (!fontProgram.TryGetGlyphId(scalar, out int glyphId) || glyphId <= 0) {
                return false;
            }
        }

        return true;
    }

    private static bool TryGetCoveredTextLength(string text, int index, PdfOpenTypeCffFontProgram fontProgram, PdfTextShapingMode shapingMode, out int length) {
        if (shapingMode == PdfTextShapingMode.LatinLigatures &&
            PdfLatinLigatureSubstitution.TryGetPresentationLigature(text, index, out int ligatureScalar, out length) &&
            fontProgram.TryGetGlyphId(ligatureScalar, out int ligatureGlyphId) &&
            ligatureGlyphId > 0) {
            return true;
        }

        int endIndex = index;
        int scalar = ReadScalar(text, ref endIndex);
        length = endIndex - index;
        return fontProgram.TryGetGlyphId(scalar, out int glyphId) && glyphId > 0;
    }

    private static bool IsLayoutControlRun(TextRun run) =>
        string.Equals(run.Text, "\n", StringComparison.Ordinal) ||
        string.Equals(run.Text, "\t", StringComparison.Ordinal);

    private static PdfAlign ResolveRichLineAlignment(PdfAlign fallback, System.Collections.Generic.IReadOnlyList<PdfAlign?>? lineAlignments, int lineIndex) =>
        lineAlignments != null && lineIndex >= 0 && lineIndex < lineAlignments.Count && lineAlignments[lineIndex].HasValue
            ? lineAlignments[lineIndex]!.Value
            : fallback;

    private static double ResolveRichLineWidth(double fallback, double? firstLineWidthOverride, System.Collections.Generic.IReadOnlyList<double>? lineWidths, int lineIndex) =>
        lineWidths != null && lineIndex >= 0 && lineIndex < lineWidths.Count
            ? lineWidths[lineIndex]
            : lineIndex == 0 ? firstLineWidthOverride ?? fallback : fallback;

    private static double ResolveRichLineXOrigin(double fallback, double? firstLineXOverride, System.Collections.Generic.IReadOnlyList<double>? lineXOffsets, int lineIndex) =>
        lineXOffsets != null && lineIndex >= 0 && lineIndex < lineXOffsets.Count
            ? fallback + lineXOffsets[lineIndex]
            : lineIndex == 0 ? firstLineXOverride ?? fallback : fallback;

    private static void WriteRichParagraph(StringBuilder sb, RichParagraphBlock block, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, PdfOptions opts, double startY, double fontSize, double defaultLeading, System.Collections.Generic.List<LinkAnnotation> annots, double? xOverride = null, double? widthOverride = null, double? firstLineXOverride = null, double? firstLineWidthOverride = null, string? structureType = null, int? markedContentId = null, LayoutResult.Page? structurePage = null, System.Collections.Generic.IReadOnlyList<PdfAlign?>? lineAlignments = null, System.Collections.Generic.IReadOnlyList<double>? lineXOffsets = null, System.Collections.Generic.IReadOnlyList<double>? lineWidths = null) {
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
            double lineWidthUsed = ResolveRichLineWidth(widthUsed, firstLineWidthOverride, lineWidths, li);
            double lineXOrigin = ResolveRichLineXOrigin(xOrigin, firstLineXOverride, lineXOffsets, li);
            var segs = lines[li];
            double baseLineW = 0;
            int gapsCount = 0;
            foreach (var seg in segs) {
                double w = MeasureRichText(seg.Text, seg.Font, seg.FontSize, seg.Baseline, opts);
                if (seg.LeadingSpace) {
                    w += seg.LeadingAdvance > 0 ? seg.LeadingAdvance : MeasureRichText(" ", seg.Font, seg.FontSize, seg.Baseline, opts);
                    if (seg.LeadingSpaceIsExpandable) {
                        gapsCount++;
                    }
                }

                baseLineW += w;
            }

            bool lineEndsWithHardBreak = segs.Any(seg => seg.EndsWithHardBreak);
            PdfAlign lineAlign = ResolveRichLineAlignment(block.Align, lineAlignments, li);
            bool justify = lineAlign == PdfAlign.Justify && !lineEndsWithHardBreak && li != lines.Count - 1 && gapsCount > 0 && lineWidthUsed > baseLineW;
            double wordSpacing = justify ? (lineWidthUsed - baseLineW) / gapsCount : 0;
            double lineWForAlign = justify ? lineWidthUsed : baseLineW;
            double dx = 0;
            if (lineAlign == PdfAlign.Center) dx = Math.Max(0, (lineWidthUsed - lineWForAlign) / 2);
            else if (lineAlign == PdfAlign.Right) dx = Math.Max(0, lineWidthUsed - lineWForAlign);

            double xCursor = dx;
            foreach (var s in segs) {
                double leadingAdvance = 0D;
                if (s.LeadingSpace) {
                    double baseGap = s.LeadingAdvance > 0 ? s.LeadingAdvance : MeasureRichText(" ", s.Font, s.FontSize, s.Baseline, opts);
                    leadingAdvance = baseGap + (s.LeadingSpaceIsExpandable ? wordSpacing : 0);
                    xCursor += leadingAdvance;
                }

                double wSeg = MeasureRichText(s.Text, s.Font, s.FontSize, s.Baseline, opts);
                if (s.BackgroundColor.HasValue && wSeg > 0) {
                    double runFontSize = EffectiveRichFontSize(s.FontSize, s.Baseline);
                    double textRise = TextRiseForBaseline(s.FontSize, s.Baseline);
                    double asc = GetAscenderForOptions(s.Font, runFontSize, opts);
                    double desc = GetDescenderForOptions(s.Font, runFontSize, opts);
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

        if (backgrounds.Count > 0) {
            AppendArtifactBegin(sb, markedContentId.HasValue);
            foreach (var bg in backgrounds) {
                new ContentStreamBuilder(sb)
                    .SaveState()
                    .FillColor(bg.Color)
                    .Rectangle(bg.X, bg.Y, bg.Width, bg.Height)
                    .FillPath()
                    .RestoreState();
            }

            AppendArtifactEnd(sb, markedContentId.HasValue);
        }

        AppendMarkedContentBegin(sb, structureType, markedContentId);
        bool textMarkedContentOpen = markedContentId.HasValue;
        int? textStructElementIndex = FindStructElementIndex(structurePage, markedContentId, structureType);
        ContentStreamBuilder content = new ContentStreamBuilder(sb)
            .BeginText()
            .TextLeading(defaultLeading);

        double yOffset = 0D;
        for (int li = 0; li < lines.Count; li++) {
            double lineY = startY - yOffset;
            double lineWidthUsed = ResolveRichLineWidth(widthUsed, firstLineWidthOverride, lineWidths, li);
            double lineXOrigin = ResolveRichLineXOrigin(xOrigin, firstLineXOverride, lineXOffsets, li);
            var segs = lines[li];
            int segCount = segs.Count;
            double[] segWidths = segCount > 0 ? new double[segCount] : System.Array.Empty<double>();
            double baseLineW = 0;
            int gapsCount = 0;
            for (int si = 0; si < segCount; si++) {
                var seg = segs[si];
                double w = MeasureRichText(seg.Text, seg.Font, seg.FontSize, seg.Baseline, opts);
                if (seg.LeadingSpace) {
                    w += seg.LeadingAdvance > 0 ? seg.LeadingAdvance : MeasureRichText(" ", seg.Font, seg.FontSize, seg.Baseline, opts);
                    if (seg.LeadingSpaceIsExpandable) {
                        gapsCount++;
                    }
                }
                segWidths[si] = w;
                baseLineW += w;
            }
            bool lineEndsWithHardBreak = segs.Any(seg => seg.EndsWithHardBreak);
            PdfAlign lineAlign = ResolveRichLineAlignment(block.Align, lineAlignments, li);
            bool justify = lineAlign == PdfAlign.Justify && !lineEndsWithHardBreak && li != lines.Count - 1 && gapsCount > 0 && lineWidthUsed > baseLineW;
            double wordSpacing = justify ? (lineWidthUsed - baseLineW) / gapsCount : 0;

            double lineWForAlign = justify ? lineWidthUsed : baseLineW;
            double dx = 0;
            if (lineAlign == PdfAlign.Center) dx = Math.Max(0, (lineWidthUsed - lineWForAlign) / 2);
            else if (lineAlign == PdfAlign.Right) dx = Math.Max(0, lineWidthUsed - lineWForAlign);
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
                bool hasLinkTarget = !string.IsNullOrEmpty(s.Uri) || !string.IsNullOrEmpty(s.DestinationName);
                if (!hasLinkTarget || s.LeadingSpace) {
                    EnsureTextMarkedContentOpen(
                        sb,
                        ref content,
                        ref textMarkedContentOpen,
                        structurePage,
                        textStructElementIndex,
                        structureType,
                        defaultLeading,
                        lineXOrigin + xCursor,
                        lineY,
                        wordSpacing,
                        fontRes,
                        runFontSize,
                        textRise,
                        color ?? PdfColor.Black);
                }

                if (s.LeadingSpace) {
                    double baseGap = s.LeadingAdvance > 0 ? s.LeadingAdvance : MeasureRichText(" ", s.Font, s.FontSize, s.Baseline, opts);
                    double gap = baseGap + (s.LeadingSpaceIsExpandable ? wordSpacing : 0);

                    if (s.LeadingTabLeader != PdfTabLeaderStyle.None) {
                        string leader = BuildTabLeaderText(gap, s.Font, s.FontSize, s.Baseline, s.LeadingTabLeader, opts);
                        if (leader.Length > 0) {
                            content
                                .TextMatrix(lineXOrigin + xCursor, lineY)
                                .ShowHexText(EncodeTextHex(leader, s.Font, opts));
                        }
                        xCursor += gap;
                        content.TextMatrix(lineXOrigin + xCursor, lineY);
                    } else if (!s.LeadingSpaceIsExpandable) {
                        content
                            .TextMatrix(lineXOrigin + xCursor, lineY)
                            .ShowHexText(EncodeTextHex(" ", s.Font, opts));
                        xCursor += gap;
                        content.TextMatrix(lineXOrigin + xCursor, lineY);
                    } else {
                        content.ShowHexText(EncodeTextHex(" ", s.Font, opts));
                        xCursor += gap;
                    }
                }
                double wSeg = MeasureRichText(s.Text, s.Font, s.FontSize, s.Baseline, opts);
                int? linkMarkedContentId = null;
                int? linkStructElementIndex = null;
                if (hasLinkTarget && opts.TaggedStructureMode == PdfTaggedStructureMode.CatalogMarkers && structurePage != null) {
                    linkMarkedContentId = structurePage.NextMarkedContentId++;
                    linkStructElementIndex = structurePage.StructElements.Count;
                    structurePage.StructElements.Add(new PageStructElement {
                        MarkedContentId = linkMarkedContentId,
                        StructureType = "Link"
                    });
                }

                double segmentStartX = xCursor;
                if (linkMarkedContentId.HasValue) {
                    content.EndText();
                    if (textMarkedContentOpen) {
                        AppendMarkedContentEnd(sb, markedContentId);
                        textMarkedContentOpen = false;
                    }

                    AppendMarkedContentBegin(sb, "Link", linkMarkedContentId);
                    content
                        .BeginText()
                        .TextLeading(defaultLeading)
                        .TextMatrix(lineXOrigin + xCursor, lineY)
                        .WordSpacing(wordSpacing)
                        .Font(fontRes, runFontSize);
                    if (Math.Abs(textRise) > 0.0001) {
                        content.TextRise(textRise);
                    }

                    content
                        .FillColor(color ?? PdfColor.Black)
                        .ShowHexText(EncodeTextHex(s.Text, s.Font, opts))
                        .EndText();
                    AppendMarkedContentEnd(sb, linkMarkedContentId);
                    content
                        .BeginText()
                        .TextLeading(defaultLeading)
                        .TextMatrix(lineXOrigin + xCursor + wSeg, lineY)
                        .WordSpacing(wordSpacing);
                    if (Math.Abs(textRise) > 0.0001) {
                        content.TextRise(0);
                    }

                    currentTextRise = 0;
                } else {
                    content.ShowHexText(EncodeTextHex(s.Text, s.Font, opts));
                }

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
                if (hasLinkTarget) {
                    var fontForMetrics = s.Font;
                    double asc = GetAscenderForOptions(fontForMetrics, runFontSize, opts);
                    double desc = GetDescenderForOptions(fontForMetrics, runFontSize, opts);
                    double x1 = lineXOrigin + segmentStartX;
                    double x2 = x1 + wSeg;
                    double y1 = baselineY - desc;
                    double y2 = baselineY + asc;
                    AddRichTextLinkAnnotation(annots, structurePage, x1, y1, x2, y2, s.Uri, s.DestinationName, s.Contents, linkStructElementIndex);
                }
                xCursor += wSeg;
            }

            if (segs.Count > 0 && segs.Any(seg => seg.EndsWithTextSeparator)) {
                RichSeg last = segs[segs.Count - 1];
                string separatorFontResource = GetStandardFontResourceName(last.Font, ChooseNormal(opts.DefaultFont));
                double separatorFontSize = EffectiveRichFontSize(last.FontSize, last.Baseline);
                double separatorTextRise = TextRiseForBaseline(last.FontSize, last.Baseline);
                content.Font(separatorFontResource, separatorFontSize);
                if (Math.Abs(separatorTextRise - currentTextRise) > 0.0001) {
                    content.TextRise(separatorTextRise);
                    currentTextRise = separatorTextRise;
                }

                content.ShowHexText(EncodeTextHex(" ", last.Font, opts));
            }

            if (Math.Abs(currentTextRise) > 0.0001) {
                content.TextRise(0);
            }

            yOffset += li < lineHeights.Count ? lineHeights[li] : defaultLeading;
        }
        content
            .WordSpacing(0)
            .EndText();
        if (textMarkedContentOpen) {
            AppendMarkedContentEnd(sb, markedContentId);
        }

        foreach (var ul in underlines) {
            AppendArtifactBegin(sb, markedContentId.HasValue);
            new ContentStreamBuilder(sb)
                .SaveState()
                .StrokeColor(ul.Color)
                .LineWidth(0.5)
                .MoveTo(ul.X1, ul.Y)
                .LineTo(ul.X2, ul.Y)
                .StrokePath()
                .RestoreState();
            AppendArtifactEnd(sb, markedContentId.HasValue);
        }
        foreach (var st in strikes) {
            AppendArtifactBegin(sb, markedContentId.HasValue);
            new ContentStreamBuilder(sb)
                .SaveState()
                .StrokeColor(st.Color)
                .LineWidth(0.5)
                .MoveTo(st.X1, st.Y)
                .LineTo(st.X2, st.Y)
                .StrokePath()
                .RestoreState();
            AppendArtifactEnd(sb, markedContentId.HasValue);
        }
    }

    private static int? FindStructElementIndex(LayoutResult.Page? structurePage, int? markedContentId, string? structureType) {
        if (structurePage == null || !markedContentId.HasValue || string.IsNullOrWhiteSpace(structureType)) {
            return null;
        }

        for (int i = 0; i < structurePage.StructElements.Count; i++) {
            PageStructElement element = structurePage.StructElements[i];
            if (element.MarkedContentId == markedContentId &&
                string.Equals(element.StructureType, structureType, StringComparison.Ordinal)) {
                return i;
            }
        }

        return null;
    }

    private static void EnsureTextMarkedContentOpen(
        StringBuilder sb,
        ref ContentStreamBuilder content,
        ref bool textMarkedContentOpen,
        LayoutResult.Page? structurePage,
        int? textStructElementIndex,
        string? structureType,
        double defaultLeading,
        double x,
        double y,
        double wordSpacing,
        string fontRes,
        double fontSize,
        double textRise,
        PdfColor fillColor) {
        if (textMarkedContentOpen ||
            structurePage == null ||
            !textStructElementIndex.HasValue ||
            string.IsNullOrWhiteSpace(structureType)) {
            return;
        }

        if (textStructElementIndex.Value < 0 || textStructElementIndex.Value >= structurePage.StructElements.Count) {
            return;
        }

        content.EndText();
        int markedContentId = structurePage.NextMarkedContentId++;
        PageStructElement element = structurePage.StructElements[textStructElementIndex.Value];
        if (element.AdditionalMarkedContentIds == null) {
            element.AdditionalMarkedContentIds = new System.Collections.Generic.List<int>();
        }

        element.AdditionalMarkedContentIds.Add(markedContentId);
        AppendMarkedContentBegin(sb, structureType, markedContentId);
        content = new ContentStreamBuilder(sb)
            .BeginText()
            .TextLeading(defaultLeading)
            .TextMatrix(x, y)
            .WordSpacing(wordSpacing)
            .Font(fontRes, fontSize)
            .FillColor(fillColor);
        if (Math.Abs(textRise) > 0.0001) {
            content.TextRise(textRise);
        }

        textMarkedContentOpen = true;
    }

    private static void AppendMarkedContentBegin(StringBuilder sb, string? structureType, int? markedContentId) {
        if (!markedContentId.HasValue || string.IsNullOrWhiteSpace(structureType)) {
            return;
        }

        sb.Append('/')
            .Append(structureType)
            .Append(" << /MCID ")
            .Append(markedContentId.Value.ToString(CultureInfo.InvariantCulture))
            .Append(" >> BDC\n");
    }

    private static void AppendMarkedContentEnd(StringBuilder sb, int? markedContentId) {
        if (markedContentId.HasValue) {
            sb.Append("EMC\n");
        }
    }

    private static void AddRichTextLinkAnnotation(System.Collections.Generic.List<LinkAnnotation> annots, LayoutResult.Page? structurePage, double x1, double y1, double x2, double y2, string? uri, string? destinationName, string? contents, int? structElementIndex) {
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
                if (structElementIndex.HasValue && previous.StructElementIndex.HasValue && structurePage != null) {
                    MergeLinkStructureElements(structurePage, previous.StructElementIndex.Value, structElementIndex.Value);
                } else if (structElementIndex.HasValue || previous.StructElementIndex.HasValue) {
                    annots.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = uri, DestinationName = destinationName, Contents = contents, StructElementIndex = structElementIndex });
                    return;
                }

                annots[annots.Count - 1] = new LinkAnnotation {
                    X1 = previous.X1,
                    Y1 = Math.Min(previous.Y1, y1),
                    X2 = Math.Max(previous.X2, x2),
                    Y2 = Math.Max(previous.Y2, y2),
                    Uri = uri,
                    DestinationName = destinationName,
                    Contents = contents,
                    StructElementIndex = previous.StructElementIndex
                };
                return;
            }
        }

        annots.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = uri, DestinationName = destinationName, Contents = contents, StructElementIndex = structElementIndex });
    }

    private static void MergeLinkStructureElements(LayoutResult.Page structurePage, int targetStructElementIndex, int mergedStructElementIndex) {
        if (targetStructElementIndex < 0 || targetStructElementIndex >= structurePage.StructElements.Count ||
            mergedStructElementIndex < 0 || mergedStructElementIndex >= structurePage.StructElements.Count ||
            targetStructElementIndex == mergedStructElementIndex) {
            return;
        }

        PageStructElement target = structurePage.StructElements[targetStructElementIndex];
        PageStructElement merged = structurePage.StructElements[mergedStructElementIndex];
        if (merged.MarkedContentId.HasValue) {
            if (target.AdditionalMarkedContentIds == null) {
                target.AdditionalMarkedContentIds = new System.Collections.Generic.List<int>();
            }

            target.AdditionalMarkedContentIds.Add(merged.MarkedContentId.Value);
        }

        if (mergedStructElementIndex == structurePage.StructElements.Count - 1) {
            structurePage.StructElements.RemoveAt(mergedStructElementIndex);
        }
    }

    private static void WriteClippedRichParagraph(StringBuilder sb, RichParagraphBlock block, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, PdfOptions opts, double startY, double fontSize, double defaultLeading, System.Collections.Generic.List<LinkAnnotation> annots, double clipX, double clipY, double clipWidth, double clipHeight, double? xOverride = null, double? widthOverride = null, double? firstLineXOverride = null, double? firstLineWidthOverride = null, string? structureType = null, int? markedContentId = null, LayoutResult.Page? structurePage = null, System.Collections.Generic.IReadOnlyList<PdfAlign?>? lineAlignments = null, System.Collections.Generic.IReadOnlyList<double>? lineXOffsets = null, System.Collections.Generic.IReadOnlyList<double>? lineWidths = null) {
        new ContentStreamBuilder(sb)
            .SaveState()
            .Rectangle(clipX, clipY, clipWidth, clipHeight)
            .ClipPath()
            .EndPath();

        WriteRichParagraph(sb, block, lines, lineHeights, opts, startY, fontSize, defaultLeading, annots, xOverride, widthOverride, firstLineXOverride, firstLineWidthOverride, structureType, markedContentId, structurePage, lineAlignments, lineXOffsets, lineWidths);

        new ContentStreamBuilder(sb)
            .RestoreState();
    }

}
