using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        // Native line-count spacing can multiply two valid Int32 values beyond the physical length range.
        // Reject malformed layout data before constructing the PDF style or starting pagination.
        private static OfficeLength NativeSpacingLength(double points, bool before) {
            try { return OfficeLength.Points(points); }
            catch (Exception error) when (error is OverflowException or ArgumentOutOfRangeException) {
                string source = before ? "Initial paragraph spacing" : "Paragraph spacing after";
                throw new InvalidDataException(source + " exceeds the supported PDF layout range.", error);
            }
        }

        private const double NativeAutomaticParagraphSpacing = 14D;

        private static bool IsNativeFirstBodyParagraph(WordParagraph paragraph) =>
            paragraph.Node().Exists && paragraph._document != null &&
            paragraph.Node() == paragraph._document.BodyNode().FirstDescendantElement(WordXml.W, "p");

        private static PdfCore.PdfParagraphStyle CreateNativeParagraphStyle(WordParagraph paragraph) =>
            CreateNativeParagraphStyle(paragraph, GetNativeDocumentDefaults(paragraph._document));

        private static PdfCore.PdfParagraphStyle CreateNativeParagraphStyle(
            WordParagraph paragraph,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap? nativeFontMap = null) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            var style = new PdfCore.PdfParagraphStyle {
                AllowFirstLineOverlapWithRunningFooter = nativeDefaults.RunningContentContext == null
            };
            double fontSize = ResolveNativeParagraphEffectiveFontSize(paragraph, nativeDefaults, styleDefaults);
            double naturalLineHeight = ResolveNativeParagraphSingleLineHeight(paragraph, nativeDefaults, styleDefaults, nativeFontMap: nativeFontMap);
            NativeLineSpacing lineSpacing = ResolveNativeParagraphLineSpacing(paragraph, styleDefaults, nativeDefaults);
            double lineHeight = lineSpacing.Resolve(fontSize, naturalLineHeight) ?? nativeDefaults.ParagraphLineHeight;
            style.FontSize = global::OfficeIMO.OfficeLength.Points(ResolveNativeParagraphLayoutFontSize(paragraph, nativeDefaults, styleDefaults));
            style.LineSpacing = lineSpacing.ToPdfLineSpacing(naturalLineHeight);
            OfficeXmlNode directSpacing = paragraph.Node().WChild("pPr").WChild("spacing");
            var spacingInLines = ResolveNativeParagraphSpacingInLines(paragraph, styleDefaults, nativeDefaults);
            NativeParagraphSpacing resolvedSpacing = ResolveNativeParagraphSpacing(paragraph, styleDefaults, nativeDefaults);
            if (resolvedSpacing.BeforeAuto == true) {
                // Automatic (HTML) spacing replaces the authored value. Word omits it
                // before the first paragraph of the document.
                style.SpacingBefore = NativeSpacingLength(IsNativeFirstBodyParagraph(paragraph) ? 0D : NativeAutomaticParagraphSpacing, before: true);
            } else if (spacingInLines.Before.HasValue) {
                style.SpacingBefore = NativeSpacingLength(spacingInLines.Before.Value, before: true);
            } else if (paragraph.LineSpacingBefore.HasValue) {
                style.SpacingBefore = paragraph.LineSpacingBefore.Value;
            } else if (GetNativeSpacingBeforePoints(directSpacing, fontSize, lineHeight) is { } directSpacingBefore) {
                style.SpacingBefore = NativeSpacingLength(directSpacingBefore, before: true);
            } else if (styleDefaults.SpacingBefore.HasValue) {
                style.SpacingBefore = NativeSpacingLength(styleDefaults.SpacingBefore.Value, before: true);
            } else if (nativeDefaults.ParagraphSpacingBeforeDeclared) {
                style.SpacingBefore = NativeSpacingLength(nativeDefaults.ParagraphSpacingBefore, before: true);
            }

            if (resolvedSpacing.AfterAuto == true) {
                style.SpacingAfter = NativeSpacingLength(NativeAutomaticParagraphSpacing, before: false);
            } else if (spacingInLines.After.HasValue) {
                style.SpacingAfter = NativeSpacingLength(spacingInLines.After.Value, before: false);
            } else if (paragraph.LineSpacingAfter.HasValue) {
                style.SpacingAfter = paragraph.LineSpacingAfter.Value;
            } else if (GetNativeSpacingAfterPoints(directSpacing, fontSize, lineHeight) is { } directSpacingAfter) {
                style.SpacingAfter = NativeSpacingLength(directSpacingAfter, before: false);
            } else if (styleDefaults.SpacingAfter.HasValue) {
                style.SpacingAfter = NativeSpacingLength(styleDefaults.SpacingAfter.Value, before: false);
            }

            if (paragraph.IndentationBefore.HasValue) {
                style.LeftIndent = paragraph.IndentationBefore.Value;
            } else if (styleDefaults.LeftIndent.HasValue) {
                style.LeftIndent = global::OfficeIMO.OfficeLength.Points(styleDefaults.LeftIndent.Value);
            }

            if (paragraph.IndentationAfter.HasValue) {
                style.RightIndent = paragraph.IndentationAfter.Value;
            } else if (styleDefaults.RightIndent.HasValue) {
                style.RightIndent = global::OfficeIMO.OfficeLength.Points(styleDefaults.RightIndent.Value);
            }

            if (paragraph.IndentationFirstLine.HasValue) {
                style.FirstLineIndent = paragraph.IndentationFirstLine.Value;
            } else if (styleDefaults.FirstLineIndent.HasValue) {
                style.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(styleDefaults.FirstLineIndent.Value);
            }

            if (paragraph.IndentationHanging.HasValue) {
                double hangingIndent = paragraph.IndentationHanging.Value.ToPoints();
                style.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(-hangingIndent);
            }

            style.LineHeight = lineHeight;
            if (nativeDefaults.DefaultTabStopWidth.HasValue) {
                style.DefaultTabStopWidth = global::OfficeIMO.OfficeLength.Points(nativeDefaults.DefaultTabStopWidth.Value);
            }

            if (resolvedSpacing.AfterAuto != true && !spacingInLines.After.HasValue && !paragraph.LineSpacingAfter.HasValue &&
                GetNativeSpacingAfterPoints(directSpacing, fontSize, lineHeight) == null &&
                !styleDefaults.SpacingAfter.HasValue) {
                style.SpacingAfter = NativeSpacingLength(nativeDefaults.ParagraphSpacingAfter, before: false);
            }

            foreach (PdfCore.PdfTabStop tabStop in ResolveNativeParagraphTabStops(
                paragraph, style.LeftIndent.ToPoints(), style.FirstLineIndent.ToPoints())) {
                style.TabStops.Add(tabStop);
            }

            style.KeepTogether = ReadNativeDirectParagraphOnOff(paragraph, "keepLines") ?? styleDefaults.KeepTogether ?? false;
            style.KeepTogetherOverflow = UsesModernNativeWordLayout(paragraph._document)
                ? PdfCore.PdfKeepTogetherOverflowBehavior.SplitAcrossPages : PdfCore.PdfKeepTogetherOverflowBehavior.SplitAcrossFrames;
            style.KeepWithNext = ReadNativeDirectParagraphOnOff(paragraph, "keepNext") ?? styleDefaults.KeepWithNext ?? false;
            style.WidowControl = ReadNativeDirectParagraphOnOff(paragraph, "widowControl") ?? styleDefaults.WidowControl ?? nativeDefaults.ParagraphWidowControl;
            ApplyNativeParagraphJustification(style, paragraph);
            return style;
        }

        private static bool ShouldSuppressNativeContextualSpacingAfter(WordParagraph paragraph, WordParagraph? nextParagraph) {
            if (nextParagraph == null ||
                paragraph.IsPageBreak ||
                nextParagraph.IsPageBreak ||
                HasNativePageBreakBefore(nextParagraph)) {
                return false;
            }

            return UsesNativeContextualSpacing(paragraph) &&
                   string.Equals(
                       GetNativeParagraphStyleIdentity(paragraph),
                       GetNativeParagraphStyleIdentity(nextParagraph),
                       StringComparison.Ordinal);
        }

        private static bool UsesNativeContextualSpacing(WordParagraph paragraph) =>
            ReadNativeDirectParagraphOnOff(paragraph, "contextualSpacing") ??
            GetNativeParagraphStyleDefaults(paragraph).ContextualSpacing ?? false;

        private static bool ShouldSuppressNativeContextualSpacingBefore(WordParagraph paragraph) {
            if (HasNativePageBreakBefore(paragraph) || paragraph.IsPageBreak) return false;
            OfficeXmlNode previous = paragraph.Node().PreviousSibling;
            if (!previous.IsW("p") || !UsesNativeContextualSpacing(paragraph)) return false;
            var previousParagraph = paragraph._document.ParagraphView(previous);
            return !previousParagraph.IsPageBreak && string.Equals(
                GetNativeParagraphStyleIdentity(paragraph), GetNativeParagraphStyleIdentity(previousParagraph), StringComparison.Ordinal);
        }

        private static string GetNativeParagraphStyleIdentity(WordParagraph paragraph) {
            IReadOnlyList<OfficeXmlNode> styleChain = GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId);
            if (styleChain.Count > 0) {
                return styleChain[styleChain.Count - 1].WAttr("styleId") ?? string.Empty;
            }

            return paragraph.StyleId ?? string.Empty;
        }

        private static bool IsNativeRenderableTextTabStop(string alignment) =>
            alignment != "bar" &&
            alignment != "clear";

        private static bool IsNativeRenderableTextTabStop(WordTabAlignment alignment) =>
            IsNativeRenderableTextTabStop(alignment.ToOpenXmlToken());

        private static double ResolveNativeParagraphLineHeight(WordParagraph paragraph, double fontSize) =>
            ResolveNativeParagraphLineHeight(paragraph, fontSize, GetNativeDocumentDefaults(paragraph._document));

        private static double ResolveNativeParagraphLineHeight(WordParagraph paragraph, double fontSize, NativeDocumentDefaults nativeDefaults) {
            return ResolveNativeParagraphLineHeight(paragraph, fontSize, nativeDefaults, GetNativeParagraphStyleDefaults(paragraph));
        }

        private static double ResolveNativeParagraphFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults, NativeParagraphStyleDefaults styleDefaults) =>
            paragraph.NativeFontSizePoints.HasValue && paragraph.NativeFontSizePoints.Value > 0
                ? paragraph.NativeFontSizePoints.Value
                : styleDefaults.FontSize ?? nativeDefaults.FontSize;

        private static double ResolveNativeParagraphEffectiveFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults, NativeParagraphStyleDefaults styleDefaults) =>
            ResolveNativeParagraphEffectiveFontSize(paragraph, nativeDefaults, styleDefaults, NativeTableRunStyleDefaults.Empty);

        private static double ResolveNativeParagraphEffectiveFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults, NativeParagraphStyleDefaults styleDefaults, NativeTableRunStyleDefaults tableRunStyleDefaults) {
            double fontSize = ResolveNativeParagraphFontSize(paragraph, nativeDefaults, styleDefaults);
            List<WordParagraph> runs = GetNativeRuns(paragraph);
            if (runs.Count == 0 && !string.IsNullOrWhiteSpace(paragraph.Text)) {
                NativeResolvedTextStyle paragraphTextStyle = ResolveNativeTextRunStyle(paragraph, tableRunStyleDefaults: tableRunStyleDefaults, nativeDefaults: nativeDefaults);
                if (paragraphTextStyle.FontSize.HasValue && paragraphTextStyle.FontSize.Value > fontSize) {
                    fontSize = paragraphTextStyle.FontSize.Value;
                }
            }

            foreach (WordParagraph run in runs) {
                if (!IsNativeParagraphMetricBearingRun(run, paragraph, nativeDefaults.NoteContentSection != null)) {
                    continue;
                }

                NativeResolvedTextStyle runTextStyle = ResolveNativeTextRunStyle(run, paragraph, tableRunStyleDefaults, nativeDefaults);
                if (runTextStyle.FontSize.HasValue && runTextStyle.FontSize.Value > fontSize) {
                    fontSize = runTextStyle.FontSize.Value;
                }
            }

            return fontSize;
        }

        private static double ResolveNativeParagraphLineHeight(
            WordParagraph paragraph,
            double fontSize,
            NativeDocumentDefaults nativeDefaults,
            NativeParagraphStyleDefaults styleDefaults,
            NativeFontMap? nativeFontMap = null) {
            double naturalLineHeight = ResolveNativeParagraphSingleLineHeight(
                paragraph,
                nativeDefaults,
                styleDefaults,
                nativeFontMap: nativeFontMap);
            return ResolveNativeParagraphLineSpacing(paragraph, styleDefaults, nativeDefaults)
                .Resolve(fontSize, naturalLineHeight) ?? nativeDefaults.ParagraphLineHeight;
        }

        private static double ResolveNativeParagraphSingleLineHeight(
            WordParagraph paragraph,
            NativeDocumentDefaults nativeDefaults,
            NativeParagraphStyleDefaults styleDefaults,
            NativeTableRunStyleDefaults tableRunStyleDefaults = default,
            NativeFontMap? nativeFontMap = null) {
            if (UsesNativeWordPerfectLineHeight(paragraph._document))
                return ResolveNativeWordPerfectLineHeight(paragraph, nativeDefaults, styleDefaults, tableRunStyleDefaults, nativeFontMap);
            List<WordParagraph> runs = GetNativeRuns(paragraph);
            double? lineHeight = null;
            foreach (WordParagraph run in runs) {
                if (!IsNativeParagraphMetricBearingRun(run, paragraph, nativeDefaults.NoteContentSection != null)) {
                    continue;
                }

                NativeCharacterStyleDefaults characterStyle =
                    GetNativeCharacterStyleDefaults(run._document, GetNativeRunProperties(run));
                lineHeight = Math.Max(
                    lineHeight ?? 0D,
                    ResolveNativeWordSingleLineHeight(
                        nativeFontMap,
                        EnumerateNativeParagraphOwnFontFamilies(run).Concat(EnumerateNativeStyleFontFamilies(
                            characterStyle, styleDefaults, tableRunStyleDefaults, nativeDefaults,
                            nativeFontMap?.UsePdfDefaultForDocumentDefaultFont != true)).ToArray()));
            }

            if (lineHeight.HasValue) return lineHeight.Value;
            OfficeXmlNode markFonts = GetNativeEmptyParagraphMarkFonts(paragraph, runs);
            return ResolveNativeWordSingleLineHeight(
                nativeFontMap,
                EnumerateNativeLatinFontFamilies(paragraph._document, markFonts).Concat(EnumerateNativeStyleFontFamilies(
                    default, styleDefaults, tableRunStyleDefaults, nativeDefaults,
                    nativeFontMap?.UsePdfDefaultForDocumentDefaultFont != true)).ToArray());
        }

        private static double ResolveNativeLineSpacingHeight(double lineSpacingPoints, string? lineSpacingRule, double fontSize, double naturalLineHeight) {
            double requestedLineHeight = lineSpacingPoints / fontSize;
            if (lineSpacingRule == "atLeast") {
                return Math.Max(naturalLineHeight, requestedLineHeight);
            }

            return requestedLineHeight;
        }

        private static PdfCore.PdfTabLeaderStyle MapNativeTabLeader(string leader) {
            if (leader == "dot" || leader == "middleDot" || leader == "heavy") {
                return PdfCore.PdfTabLeaderStyle.Dots;
            }

            if (leader == "hyphen") {
                return PdfCore.PdfTabLeaderStyle.Hyphens;
            }

            if (leader == "underscore") {
                return PdfCore.PdfTabLeaderStyle.Underscores;
            }

            return PdfCore.PdfTabLeaderStyle.None;
        }

        private static PdfCore.PdfTabLeaderStyle MapNativeTabLeader(WordTabLeader leader) =>
            MapNativeTabLeader(leader.ToOpenXmlToken());

        private static PdfCore.PdfTabAlignment MapNativeTabAlignment(string alignment) {
            if (alignment == "bar") {
                return PdfCore.PdfTabAlignment.Bar;
            }

            if (alignment == "center") {
                return PdfCore.PdfTabAlignment.Center;
            }

            if (alignment == "right") {
                return PdfCore.PdfTabAlignment.Right;
            }

            if (alignment == "decimal") {
                return PdfCore.PdfTabAlignment.DecimalSeparator;
            }

            return PdfCore.PdfTabAlignment.Left;
        }

        private static PdfCore.PdfTabAlignment MapNativeTabAlignment(WordTabAlignment alignment) =>
            MapNativeTabAlignment(alignment.ToOpenXmlToken());

        private static char? GetNativeDecimalTabSeparator(WordParagraph paragraph) {
            string? symbol = paragraph._document.Settings.DecimalSymbol;
            return symbol is { Length: 1 } && !char.IsControl(symbol[0]) && !char.IsSurrogate(symbol[0]) ? symbol[0] : null;
        }

        private static global::OfficeIMO.OfficeAlignment MapNativeParagraphAlign(string? alignment, bool allowJustify = true) {
            if (alignment == "center") {
                return global::OfficeIMO.OfficeAlignment.Center;
            }

            if (alignment == "right") {
                return global::OfficeIMO.OfficeAlignment.Right;
            }

            if (allowJustify &&
                (alignment == "both" ||
                 alignment == "distribute" ||
                 alignment == "highKashida" ||
                 alignment == "lowKashida" ||
                 alignment == "mediumKashida" ||
                 alignment == "thaiDistribute")) {
                return global::OfficeIMO.OfficeAlignment.Justify;
            }

            return global::OfficeIMO.OfficeAlignment.Left;
        }

        private static string? ResolveNativeParagraphJustification(WordParagraph paragraph) =>
            ResolveNativeVisualJustification(NativeJustificationToken(paragraph.ParagraphAlignment) ?? GetNativeParagraphStyleDefaults(paragraph).Alignment,
                IsNativeBiDiParagraph(paragraph));

        /// <summary>
        /// Word reads start and end, and also left and right, as the logical sides of the paragraph: a
        /// right-to-left paragraph starts at the right. Returns the physical side; justification is unchanged.
        /// </summary>
        private static string? ResolveNativeVisualJustification(string? alignment, bool rightToLeft) {
            if (alignment == null) return rightToLeft ? "right" : null;
            if (alignment == "start" || alignment == "left")
                return rightToLeft ? "right" : "left";
            if (alignment == "end" || alignment == "right")
                return rightToLeft ? "left" : "right";
            return alignment;
        }

        private static global::OfficeIMO.OfficeAlignment ResolveNativeParagraphAlign(WordParagraph paragraph, bool allowJustify = true) =>
            allowJustify ? MapNativeParagraphAlign(ResolveNativeParagraphJustification(paragraph))
                : MapNativeParagraphAlign(ResolveNativeUnjustifiedJustification(paragraph), allowJustify: false);

        /// <summary>
        /// The physical side for a route that sets lines without justification: justified text keeps to the
        /// paragraph's starting side, which is the right in a right-to-left paragraph.
        /// </summary>
        private static string? ResolveNativeUnjustifiedJustification(WordParagraph paragraph) =>
            ResolveNativeUnjustifiedJustification(ResolveNativeParagraphJustification(paragraph), IsNativeBiDiParagraph(paragraph));

        private static string? ResolveNativeUnjustifiedJustification(string? alignment, bool rightToLeft) =>
            MapNativeParagraphAlign(alignment) == global::OfficeIMO.OfficeAlignment.Justify
                ? rightToLeft ? "right" : "left"
                : alignment;

        private static bool IsNativeBiDiParagraph(WordParagraph paragraph) =>
            (ReadNativeDirectParagraphOnOff(paragraph, "bidi") ?? GetNativeParagraphStyleDefaults(paragraph).BiDi) == true;

        private static global::OfficeIMO.OfficeAlignment ResolveNativeColumnAlign(WordParagraph paragraph) =>
            MapNativeColumnAlign(ResolveNativeUnjustifiedJustification(paragraph));

        private static global::OfficeIMO.OfficeAlignment MapNativeColumnAlign(string? alignment) {
            if (alignment == "center") {
                return global::OfficeIMO.OfficeAlignment.Center;
            }

            if (alignment == "right") {
                return global::OfficeIMO.OfficeAlignment.Right;
            }

            return global::OfficeIMO.OfficeAlignment.Left;
        }

        private static PdfCore.PdfCellVerticalAlign MapNativeCellVerticalAlign(string? alignment) =>
            MapNativeNullableCellVerticalAlign(alignment) ?? PdfCore.PdfCellVerticalAlign.Top;

        private static PdfCore.PdfCellVerticalAlign? MapNativeNullableCellVerticalAlign(string? alignment) {
            if (alignment == null) {
                return null;
            }

            if (alignment == "center") {
                return PdfCore.PdfCellVerticalAlign.Middle;
            }

            if (alignment == "bottom") {
                return PdfCore.PdfCellVerticalAlign.Bottom;
            }

            return PdfCore.PdfCellVerticalAlign.Top;
        }

        private static PdfCore.PdfCellVerticalAlign? MapNativeNullableCellVerticalAlign(WordTableVerticalAlignment? alignment) =>
            MapNativeNullableCellVerticalAlign(NativeCellVerticalAlignmentToken(alignment));

        private static int GetHeadingLevel(WordParagraph paragraph) {
            return GetNativeTableOfContentsHeadingLevel(paragraph);
        }

        private static global::OfficeIMO.OfficeColor? GetNativeHeadingColor(int headingLevel, global::OfficeIMO.OfficeColor? explicitColor) {
            if (explicitColor.HasValue || headingLevel <= 0) {
                return explicitColor;
            }

            // Imported headings with no source color use ordinary automatic text color.
            // Declared run, style and document-default colors remain owned by run resolution.
            return global::OfficeIMO.OfficeColor.Black;
        }


        private static (double Header, double Footer) GetNativeHeaderFooterMarginExpansion(WordSection section, WordPdfExportOptions? options) {
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers =
                WordDocumentTraversal.BuildListMarkers(section._document);
            return GetNativeHeaderFooterMarginExpansion(section, options, listMarkers);
        }

        private static (double Header, double Footer) GetNativeHeaderFooterMarginExpansion(
            WordSection section,
            WordPdfExportOptions? options,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            if (options?.Margins != null) {
                return (0D, 0D);
            }

            double headerExpansion = UsesNativeRunningHeader(section, options) ? 0D : GetNativeHeaderFooterMarginExpansion(
                listMarkers,
                section.ResolveDefaultHeader(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                section.PageSetup.DifferentOddAndEvenPages ? section.ResolveEvenHeader() : null);
            double footerExpansion = UsesNativeRunningFooter(section, options) ? 0D : GetNativeFooterMarginExpansion(
                listMarkers,
                section.ResolveDefaultFooter(),
                section.PageSetup.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                section.PageSetup.DifferentOddAndEvenPages ? section.ResolveEvenFooter() : null);

            double baseTop = (checked((int?)(section.PageSetup.Margins.Top?.ToTwips())) ?? 0) / 20D;
            if (section._document.Settings.GutterAtTop) baseTop += checked((uint)(section.PageSetup.Margins.Gutter.GetValueOrDefault().ToTwips())) / 20D;
            double baseBottom = (checked((int?)(section.PageSetup.Margins.Bottom?.ToTwips())) ?? 0) / 20D;
            double availableBodyHeight = Math.Max(0D, GetNativePageSize(section, options).Height.ToPoints() - baseTop - baseBottom);
            double minimumBodyHeight = Math.Min(NativeMinimumBodyHeight, availableBodyHeight);
            double maximumExpansion = Math.Max(0D, availableBodyHeight - minimumBodyHeight);
            double totalExpansion = headerExpansion + footerExpansion;
            if (totalExpansion > maximumExpansion && totalExpansion > 0D) {
                headerExpansion = maximumExpansion * (headerExpansion / totalExpansion);
                footerExpansion = maximumExpansion - headerExpansion;
            }

            return (headerExpansion, footerExpansion);
        }

        private static double GetNativeHeaderFooterMarginExpansion(IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, params WordHeaderFooter?[] variants) {
            int maxLines = 0;
            foreach (WordHeaderFooter? variant in variants) {
                maxLines = Math.Max(maxLines, GetNativeHeaderFooterLineCount(variant, listMarkers));
            }

            return GetNativeHeaderFooterMarginExpansion(maxLines, GetNativeHeaderFooterLineHeight(variants));
        }

        private static double GetNativeFooterMarginExpansion(IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, params WordHeaderFooter?[] variants) {
            int maxLines = 0;
            foreach (WordHeaderFooter? variant in variants) {
                maxLines = Math.Max(maxLines, GetNativeHeaderFooterLineCount(variant, listMarkers));
            }

            return GetNativeFooterMarginExpansion(maxLines, GetNativeHeaderFooterLineHeight(variants));
        }

        private static double GetNativeHeaderFooterTextMarginExpansion(params NativeHeaderFooterText?[] variants) {
            int maxLines = 0;
            foreach (NativeHeaderFooterText? variant in variants) {
                maxLines = Math.Max(maxLines, GetNativeHeaderFooterTextLineCount(variant));
            }

            return GetNativeHeaderFooterMarginExpansion(maxLines);
        }

        private static double GetNativeHeaderFooterMarginExpansion(int maxLines, double lineHeight = NativeHeaderFooterLineHeight) {
            if (maxLines <= 2) {
                return 0D;
            }

            return (maxLines - 2) * lineHeight + NativeHeaderFooterBodyGap;
        }

        private static double GetNativeFooterMarginExpansion(int maxLines, double lineHeight = NativeHeaderFooterLineHeight) {
            if (maxLines <= 1) {
                return 0D;
            }

            return ((maxLines - 1) * lineHeight) + NativeHeaderFooterBodyGap;
        }

        private static double GetNativeHeaderFooterLineHeight(params WordHeaderFooter?[] variants) {
            double maxFontSize = NativeHeaderFooterFontSize;
            foreach (WordHeaderFooter? variant in variants) {
                foreach (double fontSize in EnumerateNativeHeaderFooterFontSizes(variant)) {
                    if (fontSize > maxFontSize) {
                        maxFontSize = fontSize;
                    }
                }
            }

            return maxFontSize * 1.2D;
        }

        private static int GetNativeHeaderFooterLineCount(WordHeaderFooter? headerFooter, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            if (headerFooter == null) {
                return 0;
            }

            // Some Word files carry several whitespace-only header paragraphs.
            // They paint nothing and Word does not move the body down for them.
            // Keep the existing line estimate when any visible header content exists.
            bool whitespaceOnly = true;
            foreach (WordElement element in headerFooter.Elements) {
                if (element is not WordParagraph paragraph ||
                    !string.IsNullOrWhiteSpace(GetNativeHeaderFooterParagraphText(paragraph, listMarkers, out _)) ||
                    paragraph.EnumerateImages().Any() || paragraph.PictureControl?.Image != null ||
                    paragraph.Shape != null || paragraph.Chart != null || paragraph.TextBox != null) {
                    whitespaceOnly = false;
                    break;
                }
            }
            if (whitespaceOnly) {
                return 0;
            }

            int textLines = GetNativeHeaderFooterTextLineCount(GetNativeHeaderFooterText(headerFooter, listMarkers));
            int structuralLines = 0;
            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                structuralLines += GetNativeHeaderFooterElementLineCount(element, listMarkers);
            }

            return Math.Max(textLines, structuralLines);
        }

        private static int GetNativeHeaderFooterElementLineCount(WordElement element, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            return element switch {
                WordParagraph paragraph => GetNativeHeaderFooterParagraphLineCount(paragraph, listMarkers),
                WordTable table => GetNativeHeaderFooterTableLineCount(table, listMarkers),
                WordHyperLink link when !string.IsNullOrWhiteSpace(link.Text) => 1,
                _ => 0
            };
        }

        private static int GetNativeHeaderFooterTableLineCount(WordTable table, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            int lineCount = 0;
            foreach (WordTableRow row in table.Rows) {
                int rowLineCount = 0;
                foreach (WordTableCell cell in row.Cells) {
                    int cellLineCount = 0;
                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                        cellLineCount += GetNativeHeaderFooterParagraphLineCount(paragraph, listMarkers);
                    }

                    rowLineCount = Math.Max(rowLineCount, cellLineCount);
                }

                lineCount += rowLineCount;
            }

            return lineCount;
        }

        private static int GetNativeHeaderFooterParagraphLineCount(WordParagraph paragraph, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            string? text = GetNativeHeaderFooterParagraphText(paragraph, listMarkers, out _);
            return Math.Max(1, CountNativeHeaderFooterLines(text));
        }

        private static int GetNativeHeaderFooterTextLineCount(NativeHeaderFooterText? text) {
            if (text == null) {
                return 0;
            }

            return Math.Max(
                CountNativeHeaderFooterLines(text.Left),
                Math.Max(
                    CountNativeHeaderFooterLines(text.Center),
                    CountNativeHeaderFooterLines(text.Right)));
        }

        private static int CountNativeHeaderFooterLines(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return 0;
            }

            string normalized = text!.Replace("\r\n", "\n").Replace('\r', '\n');
            int lines = 1;
            for (int index = 0; index < normalized.Length; index++) {
                if (normalized[index] == '\n') {
                    lines++;
                }
            }

            return lines;
        }

        private static string GetNativePageNumberFormat(WordPdfExportOptions? options) {
            string? format = options?.PageNumberFormat;
            if (string.IsNullOrWhiteSpace(format)) {
                return "{page}/{pages}";
            }

            return format!.Replace("{current}", "{page}").Replace("{total}", "{pages}");
        }

        private static string? BuildNativeKeywords(WordPdfExportOptions? options, WordBuiltinDocumentProperties properties) {
            return options?.Keywords ?? properties.Keywords;
        }

        private static global::OfficeIMO.OfficeColor? ParseNativeColor(string? hex) {
            if (hex == null || string.IsNullOrWhiteSpace(hex) || hex.Equals("auto", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            string value = hex.Trim();
            int metadataSeparator = value.IndexOf(' ');
            if (metadataSeparator > 0) {
                value = value.Substring(0, metadataSeparator);
            }

            if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
                !OfficeColor.TryParse(value, out OfficeColor color)) {
                return null;
            }

            return color;
        }

        private static global::OfficeIMO.OfficeColor? MapNativeHighlight(string? highlight) =>
            ParseNativeColor(WordConverterNodes.HighlightColorHexOfToken(highlight));

        /// <summary>The <c>w:jc</c> token of a paragraph alignment, or null.</summary>
        private static string? NativeJustificationToken(WordParagraphAlignment? alignment) => alignment switch {
            null => null,
            WordParagraphAlignment.Left => "left",
            WordParagraphAlignment.Start => "start",
            WordParagraphAlignment.Center => "center",
            WordParagraphAlignment.Right => "right",
            WordParagraphAlignment.End => "end",
            WordParagraphAlignment.Both => "both",
            WordParagraphAlignment.MediumKashida => "mediumKashida",
            WordParagraphAlignment.Distribute => "distribute",
            WordParagraphAlignment.NumTab => "numTab",
            WordParagraphAlignment.HighKashida => "highKashida",
            WordParagraphAlignment.LowKashida => "lowKashida",
            WordParagraphAlignment.ThaiDistribute => "thaiDistribute",
            _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, "Unsupported paragraph alignment.")
        };

        /// <summary>The <c>w:vAlign</c> token of a cell vertical alignment, or null.</summary>
        private static string? NativeCellVerticalAlignmentToken(WordTableVerticalAlignment? alignment) => alignment switch {
            null => null,
            WordTableVerticalAlignment.Top => "top",
            WordTableVerticalAlignment.Center => "center",
            WordTableVerticalAlignment.Bottom => "bottom",
            _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, "Unsupported cell vertical alignment.")
        };

    }
}
