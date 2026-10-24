using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
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
            style.FontSize = ResolveNativeParagraphLayoutFontSize(paragraph, nativeDefaults, styleDefaults);
            style.LineSpacing = lineSpacing.ToPdfLineSpacing(naturalLineHeight);
            W.SpacingBetweenLines? directSpacing = paragraph._paragraph?.ParagraphProperties?.GetFirstChild<W.SpacingBetweenLines>();
            var spacingInLines = ResolveNativeParagraphSpacingInLines(paragraph, styleDefaults, nativeDefaults);
            if (spacingInLines.Before.HasValue) {
                style.SpacingBefore = spacingInLines.Before.Value;
            } else if (paragraph.LineSpacingBeforePoints.HasValue) {
                style.SpacingBefore = paragraph.LineSpacingBeforePoints.Value;
            } else if (GetNativeSpacingBeforePoints(directSpacing, fontSize, lineHeight) is { } directSpacingBefore) {
                style.SpacingBefore = directSpacingBefore;
            } else if (styleDefaults.SpacingBefore.HasValue) {
                style.SpacingBefore = styleDefaults.SpacingBefore.Value;
            } else if (nativeDefaults.ParagraphSpacingBeforeDeclared) {
                style.SpacingBefore = nativeDefaults.ParagraphSpacingBefore;
            }

            if (spacingInLines.After.HasValue) {
                style.SpacingAfter = spacingInLines.After.Value;
            } else if (paragraph.LineSpacingAfterPoints.HasValue) {
                style.SpacingAfter = paragraph.LineSpacingAfterPoints.Value;
            } else if (GetNativeSpacingAfterPoints(directSpacing, fontSize, lineHeight) is { } directSpacingAfter) {
                style.SpacingAfter = directSpacingAfter;
            } else if (styleDefaults.SpacingAfter.HasValue) {
                style.SpacingAfter = styleDefaults.SpacingAfter.Value;
            }

            if (paragraph.IndentationBeforePoints.HasValue) {
                style.LeftIndent = paragraph.IndentationBeforePoints.Value;
            } else if (styleDefaults.LeftIndent.HasValue) {
                style.LeftIndent = styleDefaults.LeftIndent.Value;
            }

            if (paragraph.IndentationAfterPoints.HasValue) {
                style.RightIndent = paragraph.IndentationAfterPoints.Value;
            } else if (styleDefaults.RightIndent.HasValue) {
                style.RightIndent = styleDefaults.RightIndent.Value;
            }

            if (paragraph.IndentationFirstLinePoints.HasValue) {
                style.FirstLineIndent = paragraph.IndentationFirstLinePoints.Value;
            } else if (styleDefaults.FirstLineIndent.HasValue) {
                style.FirstLineIndent = styleDefaults.FirstLineIndent.Value;
            }

            if (paragraph.IndentationHangingPoints.HasValue) {
                double hangingIndent = paragraph.IndentationHangingPoints.Value;
                style.FirstLineIndent = -hangingIndent;
            }

            style.LineHeight = lineHeight;
            if (nativeDefaults.DefaultTabStopWidth.HasValue) {
                style.DefaultTabStopWidth = nativeDefaults.DefaultTabStopWidth.Value;
            }

            if (!spacingInLines.After.HasValue && !paragraph.LineSpacingAfterPoints.HasValue &&
                GetNativeSpacingAfterPoints(directSpacing, fontSize, lineHeight) == null &&
                !styleDefaults.SpacingAfter.HasValue) {
                style.SpacingAfter = nativeDefaults.ParagraphSpacingAfter;
            }

            foreach (PdfCore.PdfTabStop tabStop in ResolveNativeParagraphTabStops(
                paragraph, style.LeftIndent, style.FirstLineIndent)) {
                style.TabStops.Add(tabStop);
            }

            style.KeepTogether = ReadNativeDirectParagraphOnOff<W.KeepLines>(paragraph) ?? styleDefaults.KeepTogether ?? false;
            style.KeepTogetherOverflow = UsesModernNativeWordLayout(paragraph._document)
                ? PdfCore.PdfKeepTogetherOverflowBehavior.SplitAcrossPages : PdfCore.PdfKeepTogetherOverflowBehavior.SplitAcrossFrames;
            style.KeepWithNext = ReadNativeDirectParagraphOnOff<W.KeepNext>(paragraph) ?? styleDefaults.KeepWithNext ?? false;
            style.WidowControl = ReadNativeDirectParagraphOnOff<W.WidowControl>(paragraph) ?? styleDefaults.WidowControl ?? nativeDefaults.ParagraphWidowControl;
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
            ReadNativeDirectParagraphOnOff<W.ContextualSpacing>(paragraph) ??
            GetNativeParagraphStyleDefaults(paragraph).ContextualSpacing ?? false;

        private static bool ShouldSuppressNativeContextualSpacingBefore(WordParagraph paragraph) {
            if (HasNativePageBreakBefore(paragraph) || paragraph.IsPageBreak ||
                paragraph._paragraph?.PreviousSibling() is not W.Paragraph previous ||
                !UsesNativeContextualSpacing(paragraph)) return false;
            var previousParagraph = new WordParagraph(paragraph._document, previous);
            return !previousParagraph.IsPageBreak && string.Equals(
                GetNativeParagraphStyleIdentity(paragraph), GetNativeParagraphStyleIdentity(previousParagraph), StringComparison.Ordinal);
        }

        private static string GetNativeParagraphStyleIdentity(WordParagraph paragraph) {
            IReadOnlyList<W.Style> styleChain = GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId);
            if (styleChain.Count > 0) {
                return styleChain[styleChain.Count - 1].StyleId?.Value ?? string.Empty;
            }

            return paragraph.StyleId ?? string.Empty;
        }

        private static bool IsNativeRenderableTextTabStop(W.TabStopValues alignment) =>
            alignment != W.TabStopValues.Bar &&
            alignment != W.TabStopValues.Clear;

        private static bool IsNativeRenderableTextTabStop(WordTabAlignment alignment) =>
            IsNativeRenderableTextTabStop(alignment.ToOpenXml());

        private static double ResolveNativeParagraphLineHeight(WordParagraph paragraph, double fontSize) =>
            ResolveNativeParagraphLineHeight(paragraph, fontSize, GetNativeDocumentDefaults(paragraph._document));

        private static double ResolveNativeParagraphLineHeight(WordParagraph paragraph, double fontSize, NativeDocumentDefaults nativeDefaults) {
            return ResolveNativeParagraphLineHeight(paragraph, fontSize, nativeDefaults, GetNativeParagraphStyleDefaults(paragraph));
        }

        private static double ResolveNativeParagraphFontSize(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults, NativeParagraphStyleDefaults styleDefaults) =>
            paragraph.FontSizePoints.HasValue && paragraph.FontSizePoints.Value > 0
                ? paragraph.FontSizePoints.Value
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
            W.RunFonts? markFonts = GetNativeEmptyParagraphMarkFonts(paragraph, runs);
            return ResolveNativeWordSingleLineHeight(
                nativeFontMap,
                EnumerateNativeLatinFontFamilies(paragraph._document, markFonts).Concat(EnumerateNativeStyleFontFamilies(
                    default, styleDefaults, tableRunStyleDefaults, nativeDefaults,
                    nativeFontMap?.UsePdfDefaultForDocumentDefaultFont != true)).ToArray());
        }

        private static double ResolveNativeLineSpacingHeight(double lineSpacingPoints, W.LineSpacingRuleValues? lineSpacingRule, double fontSize, double naturalLineHeight) {
            double requestedLineHeight = lineSpacingPoints / fontSize;
            if (lineSpacingRule == W.LineSpacingRuleValues.AtLeast) {
                return Math.Max(naturalLineHeight, requestedLineHeight);
            }

            return requestedLineHeight;
        }

        private static PdfCore.PdfTabLeaderStyle MapNativeTabLeader(W.TabStopLeaderCharValues leader) {
            if (leader == W.TabStopLeaderCharValues.Dot || leader == W.TabStopLeaderCharValues.MiddleDot || leader == W.TabStopLeaderCharValues.Heavy) {
                return PdfCore.PdfTabLeaderStyle.Dots;
            }

            if (leader == W.TabStopLeaderCharValues.Hyphen) {
                return PdfCore.PdfTabLeaderStyle.Hyphens;
            }

            if (leader == W.TabStopLeaderCharValues.Underscore) {
                return PdfCore.PdfTabLeaderStyle.Underscores;
            }

            return PdfCore.PdfTabLeaderStyle.None;
        }

        private static PdfCore.PdfTabLeaderStyle MapNativeTabLeader(WordTabLeader leader) =>
            MapNativeTabLeader(leader.ToOpenXml());

        private static PdfCore.PdfTabAlignment MapNativeTabAlignment(W.TabStopValues alignment) {
            if (alignment == W.TabStopValues.Bar) {
                return PdfCore.PdfTabAlignment.Bar;
            }

            if (alignment == W.TabStopValues.Center) {
                return PdfCore.PdfTabAlignment.Center;
            }

            if (alignment == W.TabStopValues.Right) {
                return PdfCore.PdfTabAlignment.Right;
            }

            if (alignment == W.TabStopValues.Decimal) {
                return PdfCore.PdfTabAlignment.DecimalSeparator;
            }

            return PdfCore.PdfTabAlignment.Left;
        }

        private static PdfCore.PdfTabAlignment MapNativeTabAlignment(WordTabAlignment alignment) =>
            MapNativeTabAlignment(alignment.ToOpenXml());

        private static char? GetNativeDecimalTabSeparator(WordParagraph paragraph) {
            string? symbol = paragraph._document.Settings.DecimalSymbol;
            return symbol is { Length: 1 } && !char.IsControl(symbol[0]) && !char.IsSurrogate(symbol[0]) ? symbol[0] : null;
        }

        private static PdfCore.PdfAlign MapNativeParagraphAlign(W.JustificationValues? alignment, bool allowJustify = true) {
            if (alignment == W.JustificationValues.Center) {
                return PdfCore.PdfAlign.Center;
            }

            if (alignment == W.JustificationValues.Right) {
                return PdfCore.PdfAlign.Right;
            }

            if (allowJustify &&
                (alignment == W.JustificationValues.Both ||
                 alignment == W.JustificationValues.Distribute ||
                 alignment == W.JustificationValues.HighKashida ||
                 alignment == W.JustificationValues.LowKashida ||
                 alignment == W.JustificationValues.MediumKashida ||
                 alignment == W.JustificationValues.ThaiDistribute)) {
                return PdfCore.PdfAlign.Justify;
            }

            return PdfCore.PdfAlign.Left;
        }

        private static W.JustificationValues? ResolveNativeParagraphJustification(WordParagraph paragraph) =>
            paragraph.ParagraphAlignment.ToOpenXml() ?? GetNativeParagraphStyleDefaults(paragraph).Alignment;

        private static PdfCore.PdfAlign ResolveNativeParagraphAlign(WordParagraph paragraph, bool allowJustify = true) {
            W.JustificationValues? alignment = ResolveNativeParagraphJustification(paragraph);
            if (alignment == null && IsNativeBiDiParagraph(paragraph)) {
                return PdfCore.PdfAlign.Right;
            }

            return MapNativeParagraphAlign(alignment, allowJustify);
        }

        private static bool IsNativeBiDiParagraph(WordParagraph paragraph) =>
            ReadNativeDirectParagraphOnOff<W.BiDi>(paragraph) == true;

        private static PdfCore.PdfColumnAlign ResolveNativeColumnAlign(WordParagraph paragraph) =>
            MapNativeColumnAlign(ResolveNativeParagraphJustification(paragraph));

        private static PdfCore.PdfColumnAlign MapNativeColumnAlign(W.JustificationValues? alignment) {
            if (alignment == W.JustificationValues.Center) {
                return PdfCore.PdfColumnAlign.Center;
            }

            if (alignment == W.JustificationValues.Right) {
                return PdfCore.PdfColumnAlign.Right;
            }

            return PdfCore.PdfColumnAlign.Left;
        }

        private static PdfCore.PdfCellVerticalAlign MapNativeCellVerticalAlign(W.TableVerticalAlignmentValues? alignment) =>
            MapNativeNullableCellVerticalAlign(alignment) ?? PdfCore.PdfCellVerticalAlign.Top;

        private static PdfCore.PdfCellVerticalAlign? MapNativeNullableCellVerticalAlign(W.TableVerticalAlignmentValues? alignment) {
            if (!alignment.HasValue) {
                return null;
            }

            if (alignment == W.TableVerticalAlignmentValues.Center) {
                return PdfCore.PdfCellVerticalAlign.Middle;
            }

            if (alignment == W.TableVerticalAlignmentValues.Bottom) {
                return PdfCore.PdfCellVerticalAlign.Bottom;
            }

            return PdfCore.PdfCellVerticalAlign.Top;
        }

        private static PdfCore.PdfCellVerticalAlign? MapNativeNullableCellVerticalAlign(WordTableVerticalAlignment? alignment) =>
            MapNativeNullableCellVerticalAlign(alignment.ToOpenXml());

        private static int GetHeadingLevel(WordParagraph paragraph) {
            return GetNativeTableOfContentsHeadingLevel(paragraph);
        }

        private static PdfCore.PdfColor? GetNativeHeadingColor(int headingLevel, PdfCore.PdfColor? explicitColor) {
            if (explicitColor.HasValue || headingLevel <= 0) {
                return explicitColor;
            }

            return PdfCore.PdfColor.FromRgb(47, 84, 150);
        }


        private static (double Header, double Footer) GetNativeHeaderFooterMarginExpansion(WordSection section, WordToPdfOptions? options) {
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers =
                WordDocumentTraversal.BuildListMarkers(section._document);
            return GetNativeHeaderFooterMarginExpansion(section, options, listMarkers);
        }

        private static (double Header, double Footer) GetNativeHeaderFooterMarginExpansion(
            WordSection section,
            WordToPdfOptions? options,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            if (options?.Margins != null) {
                return (0D, 0D);
            }

            double headerExpansion = UsesNativeRunningHeader(section, options) ? 0D : GetNativeHeaderFooterMarginExpansion(
                listMarkers,
                section.ResolveDefaultHeader(),
                section.DifferentFirstPage ? section.ResolveFirstHeader() : null,
                section.DifferentOddAndEvenPages ? section.ResolveEvenHeader() : null);
            double footerExpansion = UsesNativeRunningFooter(section, options) ? 0D : GetNativeFooterMarginExpansion(
                listMarkers,
                section.ResolveDefaultFooter(),
                section.DifferentFirstPage ? section.ResolveFirstFooter() : null,
                section.DifferentOddAndEvenPages ? section.ResolveEvenFooter() : null);

            double baseTop = (section.Margins.Top ?? 0) / 20D;
            if (section._document.Settings.GutterAtTop) baseTop += section.Margins.Gutter / 20D;
            double baseBottom = (section.Margins.Bottom ?? 0) / 20D;
            double availableBodyHeight = Math.Max(0D, GetNativePageSize(section, options).Height - baseTop - baseBottom);
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

        private static string GetNativePageNumberFormat(WordToPdfOptions? options) {
            string? format = options?.PageNumberFormat;
            if (string.IsNullOrWhiteSpace(format)) {
                return "{page}/{pages}";
            }

            return format!.Replace("{current}", "{page}").Replace("{total}", "{pages}");
        }

        private static string? BuildNativeKeywords(WordToPdfOptions? options, WordBuiltinDocumentProperties properties) {
            return options?.Keywords ?? properties.Keywords;
        }

        private static PdfCore.PdfColor? ParseNativeColor(string? hex) {
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

            return PdfCore.PdfColor.FromOfficeColor(color);
        }

        private static PdfCore.PdfColor? MapNativeHighlight(W.HighlightColorValues? highlight) =>
            ParseNativeColor(WordRunFormatting.GetHighlightColorHex(highlight));
    }
}
