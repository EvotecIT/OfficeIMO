using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void ApplyNativeHeaderFooterPageNumberStyle(PdfCore.PdfPageBuilder page, params NativeHeaderFooterText?[] parts) {
            PdfCore.PdfPageNumberStyle? style = null;
            foreach (NativeHeaderFooterText? part in parts) {
                if (part?.PageNumberStyle == null) {
                    continue;
                }

                if (style.HasValue && style.Value != part.PageNumberStyle.Value) {
                    return;
                }

                style = part.PageNumberStyle.Value;
            }

            if (style.HasValue) {
                page.PageNumberStyle(style.Value);
            }
        }

        private static NativeHeaderFooterText? WithNativeFooterPageNumber(NativeHeaderFooterText? footer, bool includePageNumber, string pageNumberFormat) {
            if (!includePageNumber) {
                return footer;
            }

            if (footer?.HasPageTokens == true) {
                return footer;
            }

            NativeHeaderFooterText result = footer?.Clone() ?? new NativeHeaderFooterText();
            result.AppendRight(pageNumberFormat);
            return result;
        }

        private static NativeHeaderFooterText? GetNativeHeaderFooterText(WordHeaderFooter? headerFooter, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeFontMap? nativeFontMap = null) {
            if (headerFooter == null) {
                return null;
            }

            ValidateNativeHeaderFooterTextBoxNesting(headerFooter);
            var parts = new NativeHeaderFooterText();
            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                switch (element) {
                    case WordParagraph paragraph:
                        AddNativeHeaderFooterParagraphText(parts, paragraph, listMarkers, nativeFontMap: nativeFontMap);
                        break;
                    case WordTable table:
                        AddNativeHeaderFooterTableText(parts, table, listMarkers, nativeFontMap);
                        break;
                    case WordHyperLink link when !string.IsNullOrWhiteSpace(link.Text):
                        parts.AppendLeft(link.Text);
                        break;
                }
            }

            return parts.HasContent ? parts : null;
        }

        private static PdfCore.PdfStandardFont? ResolveNativeHeaderFooterFont(PdfCore.PdfStandardFont baseFont, NativeFontMap? nativeFontMap, params WordHeaderFooter?[] headerFooters) {
            PdfCore.PdfStandardFont? resolvedFont = null;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterTextStyles(headerFooter, nativeFontMap)) {
                    if (!style.Font.HasValue) {
                        continue;
                    }

                    PdfCore.PdfStandardFont fontFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(style.Font.Value);
                    if (resolvedFont.HasValue && resolvedFont.Value != fontFamily) {
                        return null;
                    }

                    resolvedFont = fontFamily;
                }
            }

            NativeHeaderFooterEmphasis emphasis = ResolveNativeHeaderFooterEmphasis(headerFooters);
            bool bold = emphasis.Bold == true;
            bool italic = emphasis.Italic == true;
            if (!resolvedFont.HasValue && !bold && !italic) {
                return null;
            }

            PdfCore.PdfStandardFont resolvedFamily = resolvedFont ?? PdfCore.PdfStandardFontMapper.GetFontFamily(baseFont);
            return PdfCore.PdfStandardFontMapper.GetStyledFont(resolvedFamily, bold, italic);
        }

        private static string? ResolveNativeHeaderFooterFontFamily(NativeFontMap nativeFontMap, params WordHeaderFooter?[] headerFooters) {
            string? resolvedFamily = null;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterTextStyles(headerFooter, nativeFontMap)) {
                    string? namedFamily = style.FontFamily;
                    if (string.IsNullOrWhiteSpace(namedFamily)) {
                        continue;
                    }

                    if (resolvedFamily != null &&
                        !string.Equals(resolvedFamily, namedFamily, StringComparison.OrdinalIgnoreCase)) {
                        return null;
                    }

                    resolvedFamily = namedFamily;
                }
            }

            return resolvedFamily;
        }

        private static PdfCore.PdfColor? ResolveNativeHeaderFooterColor(params WordHeaderFooter?[] headerFooters) {
            PdfCore.PdfColor? resolvedColor = null;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterTextStyles(headerFooter)) {
                    if (!style.Color.HasValue) {
                        return null;
                    }

                    PdfCore.PdfColor color = style.Color.Value;
                    if (resolvedColor.HasValue && !resolvedColor.Value.Equals(color)) {
                        return null;
                    }

                    resolvedColor = color;
                }
            }

            return resolvedColor;
        }

        private static double? ResolveNativeHeaderFooterFontSize(params WordHeaderFooter?[] headerFooters) {
            double? resolvedFontSize = null;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (double fontSize in EnumerateNativeHeaderFooterFontSizes(headerFooter)) {
                    if (resolvedFontSize.HasValue && !NullableDoubleEquals(resolvedFontSize.Value, fontSize)) {
                        return null;
                    }

                    resolvedFontSize = fontSize;
                }
            }

            return resolvedFontSize;
        }

        private static PdfCore.PdfStandardFont ResolveNativeHeaderFooterBaseFont(WordDocument document, WordToPdfOptions? options, bool isHeader) {
            if (options?.PdfOptions != null) {
                return PdfCore.PdfStandardFontMapper.GetFontFamily(isHeader ? options.PdfOptions.HeaderFont : options.PdfOptions.FooterFont);
            }

            foreach (string? familyName in new[] {
                options?.FontFamily,
                document.Settings.FontFamily,
                document.Settings.FontFamilyHighAnsi,
                document.Settings.FontFamilyEastAsia,
                document.Settings.FontFamilyComplexScript,
                GetNativeDocumentDefaults(document).FontFamily
            }) {
                if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfCore.PdfStandardFont mappedFont)) {
                    return PdfCore.PdfStandardFontMapper.GetFontFamily(mappedFont);
                }
            }

            return PdfCore.PdfStandardFont.Helvetica;
        }

        private readonly record struct NativeHeaderFooterEmphasis(bool? Bold, bool? Italic);

        private static NativeHeaderFooterEmphasis ResolveNativeHeaderFooterEmphasis(params WordHeaderFooter?[] headerFooters) {
            bool? bold = null;
            bool? italic = null;
            bool boldConflict = false;
            bool italicConflict = false;
            foreach (WordHeaderFooter? headerFooter in headerFooters) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterTextStyles(headerFooter)) {
                    MergeNativeHeaderFooterEmphasis(ref bold, ref boldConflict, style.Bold);
                    MergeNativeHeaderFooterEmphasis(ref italic, ref italicConflict, style.Italic);
                }
            }

            return new NativeHeaderFooterEmphasis(
                boldConflict ? null : bold,
                italicConflict ? null : italic);
        }

        private static void MergeNativeHeaderFooterEmphasis(ref bool? current, ref bool hasConflict, bool candidate) {
            if (hasConflict) {
                return;
            }

            if (!current.HasValue) {
                current = candidate;
                return;
            }

            if (current.Value != candidate) {
                current = null;
                hasConflict = true;
            }
        }

        private static IEnumerable<string> EnumerateNativeHeaderFooterFontFamilies(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                yield break;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                foreach (string familyName in EnumerateNativeHeaderFooterElementFontFamilies(element)) {
                    yield return familyName;
                }
            }
        }

        private static IEnumerable<double> EnumerateNativeHeaderFooterFontSizes(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                yield break;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                foreach (double fontSize in EnumerateNativeHeaderFooterElementFontSizes(element)) {
                    yield return fontSize;
                }
            }
        }

        private static IEnumerable<PdfCore.PdfColor> EnumerateNativeHeaderFooterColors(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                yield break;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                foreach (PdfCore.PdfColor color in EnumerateNativeHeaderFooterElementColors(element)) {
                    yield return color;
                }
            }
        }

        private static IEnumerable<NativeResolvedTextStyle> EnumerateNativeHeaderFooterTextStyles(WordHeaderFooter? headerFooter, NativeFontMap? nativeFontMap = null) {
            if (headerFooter == null) {
                yield break;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterElementTextStyles(element, nativeFontMap)) {
                    yield return style;
                }
            }
        }

        private static IEnumerable<string> EnumerateNativeHeaderFooterElementFontFamilies(WordElement element) {
            if (element is WordParagraph paragraph) {
                foreach (string familyName in EnumerateNativeParagraphFontFamilies(paragraph)) {
                    yield return familyName;
                }

                yield break;
            }

            if (element is not WordTable table) {
                yield break;
            }

            foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                foreach (WordTableRow row in currentTable.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        foreach (WordParagraph cellParagraph in cell.Paragraphs) {
                            foreach (string familyName in EnumerateNativeParagraphFontFamilies(cellParagraph)) {
                                yield return familyName;
                            }
                        }
                    }
                }
            }
        }

        private static IEnumerable<NativeResolvedTextStyle> EnumerateNativeHeaderFooterElementTextStyles(WordElement element, NativeFontMap? nativeFontMap) {
            if (element is WordParagraph paragraph) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeParagraphTextStyles(paragraph, nativeFontMap)) {
                    yield return style;
                }

                yield break;
            }

            if (element is not WordTable table) {
                yield break;
            }

            foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                foreach (WordTableRow row in currentTable.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        foreach (WordParagraph cellParagraph in cell.Paragraphs) {
                            foreach (NativeResolvedTextStyle style in EnumerateNativeParagraphTextStyles(cellParagraph, nativeFontMap)) {
                                yield return style;
                            }
                        }
                    }
                }
            }
        }

        private static IEnumerable<double> EnumerateNativeHeaderFooterElementFontSizes(WordElement element) {
            if (element is WordParagraph paragraph) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeParagraphTextStyles(paragraph)) {
                    if (style.FontSize.HasValue && style.FontSize.Value > 0D) {
                        yield return style.FontSize.Value;
                    }
                }

                yield break;
            }

            if (element is not WordTable table) {
                yield break;
            }

            foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                foreach (WordTableRow row in currentTable.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        foreach (WordParagraph cellParagraph in cell.Paragraphs) {
                            foreach (NativeResolvedTextStyle style in EnumerateNativeParagraphTextStyles(cellParagraph)) {
                                if (style.FontSize.HasValue && style.FontSize.Value > 0D) {
                                    yield return style.FontSize.Value;
                                }
                            }
                        }
                    }
                }
            }
        }

        private static IEnumerable<PdfCore.PdfColor> EnumerateNativeHeaderFooterElementColors(WordElement element) {
            if (element is WordParagraph paragraph) {
                foreach (PdfCore.PdfColor color in EnumerateNativeParagraphColors(paragraph)) {
                    yield return color;
                }

                yield break;
            }

            if (element is not WordTable table) {
                yield break;
            }

            foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                foreach (WordTableRow row in currentTable.Rows) {
                    foreach (WordTableCell cell in row.Cells) {
                        foreach (WordParagraph cellParagraph in cell.Paragraphs) {
                            foreach (PdfCore.PdfColor color in EnumerateNativeParagraphColors(cellParagraph)) {
                                yield return color;
                            }
                        }
                    }
                }
            }
        }

        private static IEnumerable<NativeResolvedTextStyle> EnumerateNativeParagraphTextStyles(WordParagraph paragraph, NativeFontMap? nativeFontMap = null) {
            List<WordParagraph> runs = GetNativeRuns(paragraph);
            bool emittedRun = false;
            foreach (WordParagraph run in runs) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) {
                    continue;
                }

                emittedRun = true;
                yield return ResolveNativeTextRunStyle(run, paragraph, nativeFontMap: nativeFontMap);
            }

            if (!emittedRun && !string.IsNullOrWhiteSpace(paragraph.Text)) {
                yield return ResolveNativeTextRunStyle(paragraph, nativeFontMap: nativeFontMap);
            }
        }

        private static IEnumerable<string> EnumerateNativeParagraphFontFamilies(WordParagraph paragraph) {
            foreach (string familyName in EnumerateNativeParagraphOwnFontFamilies(paragraph)) {
                yield return familyName;
            }

            string? styleFamily = GetNativeParagraphStyleDefaults(paragraph).FontFamily;
            if (!string.IsNullOrWhiteSpace(styleFamily)) {
                yield return styleFamily!;
            }

            string? characterStyleFamily = GetNativeCharacterStyleDefaults(paragraph._document, GetNativeRunProperties(paragraph)).FontFamily;
            if (!string.IsNullOrWhiteSpace(characterStyleFamily)) {
                yield return characterStyleFamily!;
            }

            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) {
                    continue;
                }

                foreach (string familyName in EnumerateNativeParagraphOwnFontFamilies(run)) {
                    yield return familyName;
                }

                string? runCharacterStyleFamily = GetNativeCharacterStyleDefaults(run._document, GetNativeRunProperties(run)).FontFamily;
                if (!string.IsNullOrWhiteSpace(runCharacterStyleFamily)) {
                    yield return runCharacterStyleFamily!;
                }
            }
        }

        private static IEnumerable<PdfCore.PdfColor> EnumerateNativeParagraphColors(WordParagraph paragraph) {
            PdfCore.PdfColor? paragraphColor = ParseNativeColor(paragraph.ColorHex);
            if (paragraphColor.HasValue) {
                yield return paragraphColor.Value;
            }

            PdfCore.PdfColor? styleColor = ParseNativeColor(GetNativeParagraphStyleDefaults(paragraph).ColorHex);
            if (styleColor.HasValue) {
                yield return styleColor.Value;
            }

            PdfCore.PdfColor? characterStyleColor = ParseNativeColor(GetNativeCharacterStyleDefaults(paragraph._document, GetNativeRunProperties(paragraph)).ColorHex);
            if (characterStyleColor.HasValue) {
                yield return characterStyleColor.Value;
            }

            foreach (WordParagraph run in GetNativeRuns(paragraph)) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) {
                    continue;
                }

                PdfCore.PdfColor? runColor = ParseNativeColor(run.ColorHex);
                if (runColor.HasValue) {
                    yield return runColor.Value;
                }

                PdfCore.PdfColor? runCharacterStyleColor = ParseNativeColor(GetNativeCharacterStyleDefaults(run._document, GetNativeRunProperties(run)).ColorHex);
                if (runCharacterStyleColor.HasValue) {
                    yield return runCharacterStyleColor.Value;
                }
            }
        }

        private static IEnumerable<string> EnumerateNativeParagraphOwnFontFamilies(WordParagraph paragraph) {
            foreach (string? familyName in new[] {
                paragraph.FontFamily,
                paragraph.FontFamilyHighAnsi,
                paragraph.FontFamilyEastAsia,
                paragraph.FontFamilyComplexScript
            }) {
                if (!string.IsNullOrWhiteSpace(familyName)) {
                    yield return familyName!;
                }
            }
        }

        private static IReadOnlyList<NativeHeaderFooterImage> GetNativeHeaderFooterImages(WordHeaderFooter? headerFooter, WordToPdfOptions? options, string source) {
            if (headerFooter == null) {
                return Array.Empty<NativeHeaderFooterImage>();
            }

            var images = new List<NativeHeaderFooterImage>();
            foreach (WordElement element in headerFooter.Elements) {
                switch (element) {
                    case WordParagraph paragraph:
                        AddNativeHeaderFooterParagraphImage(images, paragraph, null, options, source);
                        break;
                    case WordTable table:
                        AddNativeHeaderFooterTableImages(images, table, options, source);
                        break;
                }
            }

            return images;
        }

        private static IReadOnlyList<NativeHeaderFooterShape> GetNativeHeaderFooterShapes(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                return Array.Empty<NativeHeaderFooterShape>();
            }

            var shapes = new List<NativeHeaderFooterShape>();
            foreach (WordElement element in headerFooter.Elements) {
                switch (element) {
                    case WordParagraph paragraph:
                        AddNativeHeaderFooterParagraphShape(shapes, paragraph, null);
                        break;
                    case WordTable table:
                        AddNativeHeaderFooterTableShapes(shapes, table);
                        break;
                }
            }

            return shapes;
        }

        private static void AddNativeHeaderFooterParagraphText(NativeHeaderFooterText parts, WordParagraph paragraph, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeHeaderFooterZone? forcedZone = null, NativeFontMap? nativeFontMap = null) {
            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out _);
            string? text = GetNativeHeaderFooterParagraphText(paragraph, listMarkers, out PdfCore.PdfPageNumberStyle? pageNumberStyle, out NativeHeaderFooterZone? zoneOverride);
            IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements = textBox == null
                ? null
                : CreateNativeHeaderFooterTextBoxListReplacements(textBox, listMarkers, nativeFontMap);
            AddNativeHeaderFooterResolvedParagraphText(parts, paragraph, text, pageNumberStyle, forcedZone ?? zoneOverride, listMarkers, nativeFontMap, replacements);
        }

        private static void AddNativeHeaderFooterResolvedParagraphText(
            NativeHeaderFooterText parts,
            WordParagraph paragraph,
            string? text,
            PdfCore.PdfPageNumberStyle? pageNumberStyle,
            NativeHeaderFooterZone? resolvedZone,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            NativeFontMap? nativeFontMap,
            IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements = null) {
            string resolvedText = text ?? string.Empty;
            PdfCore.PdfTextRun? markerRun = null;
            PdfCore.PdfTextRun? contentStyleRun = null;
            WordDocumentTraversal.ListInfo? listInfo = WordDocumentTraversal.GetListInfo(paragraph);
            if (listInfo != null && listMarkers.TryGetValue(paragraph, out var marker)) {
                string serializedPrefix = marker.Marker + ResolveNativeInlineListMarkerSuffix(listInfo.Value.LevelSuffix);
                if (serializedPrefix.Length > 0 && resolvedText.StartsWith(serializedPrefix, StringComparison.Ordinal)) {
                    resolvedText = resolvedText.Substring(serializedPrefix.Length);
                }

                NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(paragraph, nativeFontMap: nativeFontMap);
                (double MarkerOffset, double TextOffset) offsets = ResolveNativeHeaderFooterListOffsets(paragraph, listInfo.Value);
                if (!string.IsNullOrEmpty(marker.Marker)) {
                    markerRun = CreateNativeHeaderFooterListMarkerTextRun(
                        marker.Marker,
                        paragraph,
                        listInfo.Value,
                        textStyle,
                        nativeFontMap,
                        offsets.MarkerOffset,
                        offsets.TextOffset);
                } else if (!string.IsNullOrEmpty(resolvedText) && offsets.TextOffset > 0D) {
                    contentStyleRun = CreateNativeHeaderFooterStyledTextRun(resolvedText, textStyle, offsets.TextOffset);
                }
            }

            if (string.IsNullOrWhiteSpace(resolvedText) && markerRun == null) {
                if (resolvedZone.HasValue) {
                    parts.Append(resolvedZone.Value, string.Empty, pageNumberStyle, markerRun: null, contentStyleRun: null);
                }
                return;
            }

            if (resolvedZone.HasValue) {
                parts.Append(resolvedZone.Value, resolvedText, pageNumberStyle, markerRun, contentStyleRun, replacements);
                return;
            }

            W.JustificationValues? alignment = ResolveNativeParagraphJustification(paragraph);
            if (alignment == W.JustificationValues.Center) {
                parts.AppendCenter(resolvedText, pageNumberStyle, markerRun, contentStyleRun, replacements);
            } else if (alignment == W.JustificationValues.Right) {
                parts.AppendRight(resolvedText, pageNumberStyle, markerRun, contentStyleRun, replacements);
            } else {
                parts.AppendLeft(resolvedText, pageNumberStyle, markerRun, contentStyleRun, replacements);
            }
        }

        private static IReadOnlyList<NativeHeaderFooterStyledReplacement> CreateNativeHeaderFooterTextBoxListReplacements(
            WordTextBox textBox,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            NativeFontMap? nativeFontMap) {
            var replacements = new List<NativeHeaderFooterStyledReplacement>();
            var pending = new Stack<WordTextBox>();
            pending.Push(textBox);
            while (pending.Count > 0) {
                WordTextBox current = pending.Pop();
                IReadOnlyList<WordParagraph> paragraphs = GetNativeTextBoxParagraphs(current);
                for (int index = paragraphs.Count - 1; index >= 0; index--) {
                    WordTextBox? nested = GetNativeParagraphTextBox(paragraphs[index], out _);
                    if (nested != null) {
                        pending.Push(nested);
                    }
                }

                foreach (WordParagraph innerParagraph in paragraphs) {
                    WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(innerParagraph);
                    if (info == null ||
                        !listMarkers.TryGetValue(innerParagraph, out var marker) ||
                        string.IsNullOrEmpty(marker.Marker)) {
                        continue;
                    }

                    string serializedPrefix = marker.Marker + ResolveNativeInlineListMarkerSuffix(info.Value.LevelSuffix);
                    NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(innerParagraph, nativeFontMap: nativeFontMap);
                    (double MarkerOffset, double TextOffset) offsets = ResolveNativeHeaderFooterListOffsets(innerParagraph, info.Value);
                    PdfCore.PdfTextRun styledMarker = CreateNativeHeaderFooterListMarkerTextRun(
                        marker.Marker,
                        innerParagraph,
                        info.Value,
                        textStyle,
                        nativeFontMap,
                        offsets.MarkerOffset,
                        offsets.TextOffset);
                    replacements.Add(new NativeHeaderFooterStyledReplacement(serializedPrefix, styledMarker));
                }
            }
            return replacements;
        }

        private static (double MarkerOffset, double TextOffset) ResolveNativeHeaderFooterListOffsets(
            WordParagraph paragraph,
            WordDocumentTraversal.ListInfo info) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            bool useParagraphStyleIndent = ShouldApplyNativeListParagraphStyleIndent(paragraph);
            double textOffset = paragraph.IndentationBeforePoints ??
                (useParagraphStyleIndent ? styleDefaults.LeftIndent : null) ??
                ConvertNativeTwipsToPoints(info.LeftIndentTwips ?? ((info.Level + 1) * 720)) ?? 0D;
            double hangingIndent = paragraph.IndentationHangingPoints ??
                (useParagraphStyleIndent ? GetNativeStyleHangingIndent(styleDefaults) : null) ??
                ConvertNativeTwipsToPoints(info.HangingIndentTwips ?? 360) ?? 0D;
            textOffset = Math.Max(0D, textOffset);
            return (Math.Max(0D, textOffset - Math.Max(0D, hangingIndent)), textOffset);
        }

        private static PdfCore.PdfTextRun CreateNativeHeaderFooterListMarkerTextRun(
            string marker,
            WordParagraph paragraph,
            WordDocumentTraversal.ListInfo info,
            NativeResolvedTextStyle textStyle,
            NativeFontMap? nativeFontMap,
            double markerOffset,
            double textOffset) {
            PdfCore.PdfTextRun styledMarker = CreateNativeListMarkerTextRun(marker, paragraph, textStyle, nativeFontMap);
            double markerFontSize = styledMarker.FontSize ?? textStyle.FontSize ?? 12D;
            string suffix = ResolveNativeHeaderFooterListMarkerSuffix(info.LevelSuffix, marker, markerFontSize, markerOffset, textOffset);
            return CloneNativeHeaderFooterTextRun(styledMarker, marker + suffix)
                .WithHorizontalOffset(markerOffset);
        }

        private static string ResolveNativeHeaderFooterListMarkerSuffix(
            WordListLevelSuffix? suffix,
            string marker,
            double markerFontSize,
            double markerOffset,
            double textOffset) {
            if (suffix == WordListLevelSuffix.Nothing) {
                return string.Empty;
            }
            if (suffix == WordListLevelSuffix.Space) {
                return " ";
            }

            double desiredGap = Math.Max(0D, textOffset - markerOffset - EstimateNativeListMarkerWidth(marker, markerFontSize));
            if (desiredGap <= 0D) {
                return string.Empty;
            }
            double spaceWidth = Math.Max(0.01D, EstimateNativeListMarkerWidth(" ", markerFontSize));
            return new string(' ', Math.Max(1, (int)Math.Ceiling(desiredGap / spaceWidth)));
        }

        private static PdfCore.PdfTextRun CreateNativeHeaderFooterStyledTextRun(
            string text,
            NativeResolvedTextStyle style,
            double horizontalOffset) =>
            new PdfCore.PdfTextRun(
                text.Replace('\t', ' '),
                bold: style.Bold,
                color: style.Color,
                italic: style.Italic,
                fontSize: style.FontSize,
                font: style.Font,
                fontFamily: style.FontFamily)
            .WithHorizontalOffset(horizontalOffset);

        private static PdfCore.PdfTextRun CloneNativeHeaderFooterTextRun(PdfCore.PdfTextRun source, string text) =>
            new PdfCore.PdfTextRun(
                text.Replace('\t', ' '),
                source.Bold,
                source.Underline,
                source.Color,
                source.Italic,
                source.Strike,
                source.FontSize,
                source.Font,
                baseline: source.Baseline,
                backgroundColor: source.BackgroundColor,
                fontFamily: source.FontFamily,
                underlineStyle: source.UnderlineStyle,
                strikeStyle: source.StrikeStyle,
                decorationColor: source.DecorationColor)
            .WithFeatureSettings(source.FeatureSettings)
            .WithHorizontalOffset(source.HorizontalOffset);

        private static void AddNativeHeaderFooterTableImages(List<NativeHeaderFooterImage> images, WordTable table, WordToPdfOptions? options, string source) {
            foreach (WordTableRow row in table.Rows) {
                IReadOnlyList<WordTableCell> cells = row.Cells;
                if (cells.Count == 1) {
                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cells[0])) {
                        AddNativeHeaderFooterParagraphImage(images, paragraph, null, options, source);
                    }

                    continue;
                }

                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    PdfCore.PdfAlign align = cellIndex == 0
                        ? PdfCore.PdfAlign.Left
                        : cellIndex == cells.Count - 1
                            ? PdfCore.PdfAlign.Right
                            : PdfCore.PdfAlign.Center;

                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cells[cellIndex])) {
                        AddNativeHeaderFooterParagraphImage(images, paragraph, align, options, source);
                    }
                }
            }
        }

        private static void AddNativeHeaderFooterParagraphImage(List<NativeHeaderFooterImage> images, WordParagraph paragraph, PdfCore.PdfAlign? alignOverride, WordToPdfOptions? options, string source) {
            PdfCore.PdfAlign align = alignOverride ?? ResolveNativeParagraphAlign(paragraph, allowJustify: false);
            if (paragraph.Image != null) {
                AddNativeHeaderFooterImage(images, paragraph.Image, align, options, source);
            }

            foreach (W.SdtRun pictureControl in GetNativePictureControls(paragraph)) {
                var pictureParagraph = new WordParagraph(paragraph._document, paragraph._paragraph!, pictureControl);
                WordImage? pictureControlImage = pictureParagraph.PictureControl?.Image;
                if (pictureControlImage == null) {
                    continue;
                }

                AddNativeHeaderFooterImage(images, pictureControlImage, align, options, source);
            }
        }

        private static void AddNativeHeaderFooterImage(List<NativeHeaderFooterImage> images, WordImage image, PdfCore.PdfAlign align, WordToPdfOptions? options, string source) {
            byte[] bytes = ImageEmbedder.GetImageBytes(image);
            if (!TryPrepareNativePdfImageBytes(bytes, out byte[] preparedBytes, out string? unsupportedReason)) {
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "NativeHeaderFooterImageUnsupported",
                        source,
                        "Word header/footer image was not exported because the shared PDF raster pipeline could not prepare it. " + unsupportedReason);
                }

                return;
            }

            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            images.Add(new NativeHeaderFooterImage(preparedBytes, width, height, align));
        }

        private static void AddNativeHeaderFooterTableShapes(List<NativeHeaderFooterShape> shapes, WordTable table) {
            foreach (WordTableRow row in table.Rows) {
                IReadOnlyList<WordTableCell> cells = row.Cells;
                if (cells.Count == 1) {
                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cells[0])) {
                        AddNativeHeaderFooterParagraphShape(shapes, paragraph, null);
                    }

                    continue;
                }

                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    PdfCore.PdfAlign align = cellIndex == 0
                        ? PdfCore.PdfAlign.Left
                        : cellIndex == cells.Count - 1
                            ? PdfCore.PdfAlign.Right
                            : PdfCore.PdfAlign.Center;

                    foreach (WordParagraph paragraph in GetNativeCellParagraphs(cells[cellIndex])) {
                        AddNativeHeaderFooterParagraphShape(shapes, paragraph, align);
                    }
                }
            }
        }

        private static void AddNativeHeaderFooterParagraphShape(List<NativeHeaderFooterShape> shapes, WordParagraph paragraph, PdfCore.PdfAlign? alignOverride) {
            if (paragraph.Shape == null) {
                return;
            }

            OfficeShape? shape = CreateNativeShape(paragraph.Shape);
            if (shape == null) {
                return;
            }

            PdfCore.PdfAlign align = alignOverride ?? ResolveNativeParagraphAlign(paragraph, allowJustify: false);
            shapes.Add(new NativeHeaderFooterShape(shape, align));
        }

        private static string? GetNativeHeaderFooterParagraphText(
            WordParagraph paragraph,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            out PdfCore.PdfPageNumberStyle? pageNumberStyle,
            int textBoxDepth = 0) {
            return GetNativeHeaderFooterParagraphText(paragraph, listMarkers, out pageNumberStyle, out _, textBoxDepth);
        }

        private static string? GetNativeHeaderFooterParagraphText(
            WordParagraph paragraph,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            out PdfCore.PdfPageNumberStyle? pageNumberStyle,
            out NativeHeaderFooterZone? zoneOverride,
            int textBoxDepth = 0) {
            zoneOverride = null;
            WordTextBox? numberedTextBox = GetNativeParagraphTextBox(paragraph, out _);
            if (numberedTextBox != null && GetNativeTextBoxParagraphs(numberedTextBox).Any(inner => inner.IsListItem)) {
                zoneOverride = MapNativeTextBoxHeaderFooterZone(numberedTextBox.HorizontalAlignment);
            }
            if (TryBuildNativeHeaderFooterParagraphText(paragraph, out string? mixedText, out pageNumberStyle)) {
                return PrependNativeHeaderFooterListMarker(paragraph, AppendNativeHeaderFooterSupplementalText(mixedText, paragraph), listMarkers, textBoxDepth);
            }

            if (TryGetNativeHeaderFooterFieldToken(paragraph, out string? fieldToken, out pageNumberStyle)) {
                return PrependNativeHeaderFooterListMarker(paragraph, AppendNativeHeaderFooterSupplementalText(fieldToken, paragraph), listMarkers, textBoxDepth);
            }

            pageNumberStyle = null;
            if (paragraph.IsHyperLink && paragraph.Hyperlink != null && !IsNativeHiddenTextRun(paragraph)) {
                return PrependNativeHeaderFooterListMarker(paragraph, AppendNativeHeaderFooterSupplementalText(ApplyNativeTextTransform(paragraph.Hyperlink.Text, paragraph), paragraph), listMarkers, textBoxDepth);
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            string? text = runs.Count > 0
                ? string.Concat(runs.Where(run => !IsNativeHiddenTextRun(run, paragraph)).Select(run => ApplyNativeTextTransform(run.Text, run, paragraph)))
                : IsNativeHiddenTextRun(paragraph) ? string.Empty : ApplyNativeTextTransform(paragraph.Text, paragraph);
            text = AppendNativeHeaderFooterSupplementalText(text, paragraph);
            if (!string.IsNullOrWhiteSpace(text)) {
                return PrependNativeHeaderFooterListMarker(paragraph, text, listMarkers, textBoxDepth);
            }

            string? textBoxText = GetNativeParagraphTextBoxPlainText(paragraph);
            if (string.IsNullOrWhiteSpace(textBoxText)) {
                return text;
            }

            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out _);
            zoneOverride = MapNativeTextBoxHeaderFooterZone(textBox?.HorizontalAlignment ?? WordTextBoxHorizontalAlignment.Center);
            if (textBox != null) {
                EnsureNativeHeaderFooterTextBoxDepth(textBoxDepth);
                string innerText = string.Join("\n", GetNativeTextBoxParagraphs(textBox)
                    .Select(inner => GetNativeHeaderFooterParagraphText(inner, listMarkers, out _, textBoxDepth + 1))
                    .Where(inner => !string.IsNullOrWhiteSpace(inner)));
                if (!string.IsNullOrWhiteSpace(innerText)) {
                    return PrependNativeHeaderFooterListMarker(paragraph, innerText, listMarkers, textBoxDepth);
                }
            }
            return PrependNativeHeaderFooterListMarker(paragraph, textBoxText, listMarkers, textBoxDepth);
        }

        private static string? PrependNativeHeaderFooterListMarker(
            WordParagraph paragraph,
            string? text,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            int textBoxDepth) {
            if (string.IsNullOrWhiteSpace(text)) return text;
            string resolvedText = text!;
            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out _);
            if (textBox != null) {
                IReadOnlyList<WordParagraph> innerParagraphs = GetNativeTextBoxParagraphs(textBox);
                if (innerParagraphs.Any(inner => inner.IsListItem)) {
                    EnsureNativeHeaderFooterTextBoxDepth(textBoxDepth);
                    string? flattenedText = GetNativeParagraphTextBoxPlainText(paragraph);
                    int textBoxStart = string.IsNullOrEmpty(flattenedText)
                        ? -1
                        : resolvedText.IndexOf(flattenedText, StringComparison.Ordinal);
                    if (textBoxStart >= 0) {
                        string innerText = string.Join("\n", innerParagraphs
                            .Select(inner => GetNativeHeaderFooterParagraphText(inner, listMarkers, out _, textBoxDepth + 1))
                            .Where(inner => !string.IsNullOrWhiteSpace(inner)));
                        resolvedText = resolvedText.Remove(textBoxStart, flattenedText!.Length).Insert(textBoxStart, innerText);
                    }
                }
            }
            if (!paragraph.IsListItem) return resolvedText;
            return listMarkers.TryGetValue(paragraph, out var marker) && !string.IsNullOrEmpty(marker.Marker)
                ? marker.Marker + ResolveNativeInlineListMarkerSuffix(WordDocumentTraversal.GetListInfo(paragraph)?.LevelSuffix) + resolvedText
                : resolvedText;
        }

        private static void EnsureNativeHeaderFooterTextBoxDepth(int textBoxDepth) {
            if (textBoxDepth >= MaximumNativeHeaderFooterTextBoxNestingDepth) {
                throw new InvalidDataException(
                    $"Header or footer text-box nesting exceeds the supported limit of {MaximumNativeHeaderFooterTextBoxNestingDepth} levels.");
            }
        }

        private static string? AppendNativeHeaderFooterSupplementalText(string? text, WordParagraph paragraph) {
            text = AppendNativeHeaderFooterEquationText(text, paragraph);
            text = AppendNativeHeaderFooterFormControlText(text, paragraph);
            text = AppendNativeHeaderFooterTextPathText(text, paragraph);
            return AppendNativeHeaderFooterRepeatingSectionText(text, paragraph);
        }

        private static string? AppendNativeHeaderFooterEquationText(string? text, WordParagraph paragraph) {
            IReadOnlyList<WordEquationOccurrence> occurrences = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph);
            if (occurrences.Count > 0) {
                string orderedText = AppendNativeTextWithEquation(text ?? string.Empty, paragraph);
                if (!string.IsNullOrWhiteSpace(orderedText)) {
                    return orderedText;
                }
            }

            string? equationText = GetNativeEquationText(paragraph);
            if (string.IsNullOrWhiteSpace(equationText)) {
                return text;
            }

            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            AppendNativeHeaderFooterSupplementalValue(builder, ref currentText, equationText, skipIfPresent: true);
            return builder.Length == 0 ? text : builder.ToString();
        }

        private static string? AppendNativeHeaderFooterFormControlText(string? text, WordParagraph paragraph) {
            IReadOnlyList<W.SdtRun> checkBoxes = GetNativeCheckBoxControls(paragraph);
            IReadOnlyList<W.SdtRun> formFields = GetNativeFormFieldControls(paragraph);
            if (checkBoxes.Count == 0 && formFields.Count == 0) {
                return text;
            }

            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            foreach (W.SdtRun checkBox in checkBoxes) {
                AppendNativeHeaderFooterSupplementalValue(
                    builder,
                    ref currentText,
                    IsNativeCheckBoxChecked(checkBox) ? "[x]" : "[ ]",
                    skipIfPresent: false);
            }

            foreach (W.SdtRun formField in formFields) {
                string? value;
                if (IsNativeDatePickerControl(formField)) {
                    value = GetNativeDatePickerValue(formField);
                } else {
                    IReadOnlyList<string> options = GetNativeChoiceFieldOptions(formField);
                    value = GetNativeChoiceFieldValue(formField, options);
                }

                AppendNativeHeaderFooterSupplementalValue(builder, ref currentText, value, skipIfPresent: true);
            }

            return builder.Length == 0 ? text : builder.ToString();
        }

        private static string? AppendNativeHeaderFooterTextPathText(string? text, WordParagraph paragraph) {
            IEnumerable<V.TextPath> textPaths = paragraph._paragraph?.Descendants<V.TextPath>() ?? Enumerable.Empty<V.TextPath>();
            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            foreach (V.TextPath textPath in textPaths) {
                if (!IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(textPath, "on")) ||
                    IsNativeHeaderFooterWatermarkTextPath(textPath)) {
                    continue;
                }

                AppendNativeHeaderFooterSupplementalValue(
                    builder,
                    ref currentText,
                    textPath.String?.Value ?? GetNativeOpenXmlAttribute(textPath, "string"),
                    skipIfPresent: true);
            }

            return builder.Length == 0 ? text : builder.ToString();
        }

        private static bool IsNativeHeaderFooterWatermarkTextPath(V.TextPath textPath) {
            V.Shape? shape = textPath.Ancestors<V.Shape>().FirstOrDefault();
            if (shape == null) {
                return false;
            }

            string marker = string.Join(" ",
                shape.Id?.Value,
                GetNativeOpenXmlAttribute(shape, "name"),
                GetNativeOpenXmlAttribute(shape, "title"));
            return marker.IndexOf("watermark", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string? AppendNativeHeaderFooterRepeatingSectionText(string? text, WordParagraph paragraph) {
            IReadOnlyList<W.SdtRun> controls = GetNativeRepeatingSectionControls(paragraph);
            if (controls.Count == 0) {
                return text;
            }

            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            foreach (W.SdtRun control in controls) {
                foreach (string itemText in GetNativeRepeatingSectionItems(control)) {
                    AppendNativeHeaderFooterSupplementalValue(builder, ref currentText, itemText, skipIfPresent: true);
                }
            }

            return builder.Length == 0 ? text : builder.ToString();
        }

        private static void AppendNativeHeaderFooterSupplementalValue(StringBuilder builder, ref string currentText, string? value, bool skipIfPresent) {
            if (string.IsNullOrWhiteSpace(value) ||
                skipIfPresent && currentText.IndexOf(value!, StringComparison.Ordinal) >= 0) {
                return;
            }

            if (builder.Length > 0 && !char.IsWhiteSpace(builder[builder.Length - 1])) {
                builder.Append(' ');
            }

            builder.Append(value);
            currentText = builder.ToString();
        }

        private static string AppendNativeTextWithEquation(string text, WordParagraph paragraph) {
            IReadOnlyList<WordEquationContentSegment> segments = GetNativeVisibleEquationContentSegments(paragraph);
            if (segments.Count == 0) {
                return text;
            }

            string orderedText = string.Concat(segments.Select(GetNativeEquationSegmentText));
            return string.IsNullOrEmpty(orderedText) ? text : orderedText;
        }

        private static IReadOnlyList<WordEquationContentSegment> GetNativeVisibleEquationContentSegments(WordParagraph paragraph) {
            IReadOnlyList<WordEquationOccurrence> occurrences = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph);
            if (occurrences.Count == 0 || GetNativeParagraphStyleDefaults(paragraph).Hidden == true) {
                return Array.Empty<WordEquationContentSegment>();
            }

            return WordEquation.GetVisibleContentSegments(
                paragraph._paragraph,
                occurrences,
                element => element is not W.Run run ||
                    !IsNativeHiddenTextRun(new WordParagraph(paragraph._document, paragraph._paragraph, run), paragraph));
        }

        private static string? GetNativeEquationText(WordParagraph paragraph) {
            string[] equationTexts = WordEquation
                .GetOccurrences(paragraph._document, paragraph._paragraph)
                .Select(occurrence => occurrence.Equation.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToArray();
            if (equationTexts.Length > 0) {
                return string.Join(" ", equationTexts);
            }

            string ommlText = WordMath.GetText(paragraph._paragraph);
            return string.IsNullOrWhiteSpace(ommlText) ? null : ommlText;
        }

        private static NativeHeaderFooterZone MapNativeTextBoxHeaderFooterZone(WordTextBoxHorizontalAlignment alignment) {
            switch (alignment) {
                case WordTextBoxHorizontalAlignment.Center:
                    return NativeHeaderFooterZone.Center;
                case WordTextBoxHorizontalAlignment.Right:
                case WordTextBoxHorizontalAlignment.Outside:
                    return NativeHeaderFooterZone.Right;
                default:
                    return NativeHeaderFooterZone.Left;
            }
        }

        private static bool TryBuildNativeHeaderFooterParagraphText(WordParagraph paragraph, out string? text, out PdfCore.PdfPageNumberStyle? pageNumberStyle) {
            text = null;
            pageNumberStyle = null;
            if (paragraph._paragraph == null) {
                return false;
            }

            var builder = new StringBuilder();
            var state = new NativeHeaderFooterFieldState();
            bool hasFieldToken = false;
            bool hasConflictingStyles = false;
            foreach (var element in paragraph._paragraph.ChildElements) {
                AppendNativeHeaderFooterElementText(element, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
            }

            if (!hasFieldToken) {
                pageNumberStyle = null;
                return false;
            }

            text = builder.ToString();
            return !string.IsNullOrWhiteSpace(text);
        }

        private static void AppendNativeHeaderFooterElementText(DocumentFormat.OpenXml.OpenXmlElement element, StringBuilder builder, NativeHeaderFooterFieldState state, ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles, ref bool hasFieldToken) {
            if (element is W.Run run) {
                AppendNativeHeaderFooterRunText(run, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                return;
            }

            if (element is W.Hyperlink hyperlink) {
                foreach (W.Run childRun in hyperlink.Elements<W.Run>()) {
                    AppendNativeHeaderFooterRunText(childRun, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }

                return;
            }

            if (element is W.SdtRun sdtRun) {
                foreach (var child in sdtRun.SdtContentRun?.ChildElements ?? Enumerable.Empty<DocumentFormat.OpenXml.OpenXmlElement>()) {
                    AppendNativeHeaderFooterElementText(child, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }

                return;
            }

            if (element is W.SimpleField simpleField) {
                string fieldCode = simpleField.Instruction?.Value ?? string.Empty;
                if (TryGetNativeHeaderFooterFieldToken(fieldCode, out string? token, out PdfCore.PdfPageNumberStyle? style)) {
                    builder.Append(token);
                    MergeNativeHeaderFooterPageNumberStyle(ref pageNumberStyle, ref hasConflictingStyles, style);
                    hasFieldToken = true;
                    return;
                }

                foreach (var child in simpleField.ChildElements) {
                    AppendNativeHeaderFooterElementText(child, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }
            }
        }

        private static void AppendNativeHeaderFooterRunText(W.Run run, StringBuilder builder, NativeHeaderFooterFieldState state, ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles, ref bool hasFieldToken) {
            if (IsNativeHiddenRun(run)) {
                return;
            }

            foreach (var child in run.ChildElements) {
                if (child is W.FieldChar fieldChar) {
                    W.FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                    if (fieldCharType == W.FieldCharValues.Begin) {
                        state.CollectingFieldCode = true;
                        state.SkippingFieldResult = false;
                        state.FieldCode.Clear();
                    } else if (fieldCharType == W.FieldCharValues.Separate) {
                        if (TryGetNativeHeaderFooterFieldToken(state.FieldCode.ToString(), out string? token, out PdfCore.PdfPageNumberStyle? style)) {
                            builder.Append(token);
                            MergeNativeHeaderFooterPageNumberStyle(ref pageNumberStyle, ref hasConflictingStyles, style);
                            hasFieldToken = true;
                            state.SkippingFieldResult = true;
                        }

                        state.CollectingFieldCode = false;
                    } else if (fieldCharType == W.FieldCharValues.End) {
                        state.CollectingFieldCode = false;
                        state.SkippingFieldResult = false;
                        state.FieldCode.Clear();
                    }

                    continue;
                }

                if (child is W.FieldCode fieldCode) {
                    if (state.CollectingFieldCode) {
                        state.FieldCode.Append(fieldCode.Text);
                    }

                    continue;
                }

                if (state.CollectingFieldCode || state.SkippingFieldResult) {
                    continue;
                }

                if (child is W.Text text) {
                    builder.Append(ApplyNativeHeaderFooterRunTextTransform(text.Text, run));
                } else if (child is W.TabChar) {
                    builder.Append('\t');
                } else if (child is W.Break) {
                    builder.AppendLine();
                }
            }
        }

        private static bool IsNativeHiddenRun(W.Run run) =>
            ReadNativeOnOff(run.RunProperties?.GetFirstChild<W.Vanish>()) == true;

        private static string ApplyNativeHeaderFooterRunTextTransform(string text, W.Run run) =>
            IsNativeAllCapsRun(run)
                ? text.ToUpperInvariant()
                : text;

        private static bool IsNativeAllCapsRun(W.Run run) =>
            ReadNativeOnOff(run.RunProperties?.GetFirstChild<W.Caps>()) == true ||
            ReadNativeOnOff(run.RunProperties?.GetFirstChild<W.SmallCaps>()) == true;

        private static bool TryGetNativeHeaderFooterFieldToken(WordParagraph paragraph, out string? token, out PdfCore.PdfPageNumberStyle? style) {
            token = null;
            style = null;
            WordField? field = paragraph.Field;
            if (field?.FieldType == WordFieldType.Page) {
                token = "{page}";
                style = MapNativePageNumberFieldStyle(field.Field);
                return true;
            }

            if (field?.FieldType == WordFieldType.NumPages) {
                token = "{documentpages}";
                style = MapNativePageNumberFieldStyle(field.Field);
                return true;
            }

            if (field?.FieldType == WordFieldType.SectionPages) {
                token = "{pages}";
                style = MapNativePageNumberFieldStyle(field.Field);
                return true;
            }

            return false;
        }

        private static bool TryGetNativeHeaderFooterFieldToken(string fieldCode, out string? token, out PdfCore.PdfPageNumberStyle? style) {
            token = null;
            style = null;
            string trimmed = fieldCode.Trim();
            if (trimmed.Length == 0) {
                return false;
            }

            int end = 0;
            while (end < trimmed.Length && !char.IsWhiteSpace(trimmed[end])) {
                end++;
            }

            string fieldType = trimmed.Substring(0, end);
            if (string.Equals(fieldType, "PAGE", StringComparison.OrdinalIgnoreCase)) {
                token = "{page}";
                style = MapNativePageNumberFieldStyle(trimmed);
                return true;
            }

            if (string.Equals(fieldType, "NUMPAGES", StringComparison.OrdinalIgnoreCase)) {
                token = "{documentpages}";
                style = MapNativePageNumberFieldStyle(trimmed);
                return true;
            }

            if (string.Equals(fieldType, "SECTIONPAGES", StringComparison.OrdinalIgnoreCase)) {
                token = "{pages}";
                style = MapNativePageNumberFieldStyle(trimmed);
                return true;
            }

            return false;
        }

        private static PdfCore.PdfPageNumberStyle? MapNativePageNumberFieldStyle(string fieldCode) {
            string? format = GetNativePageNumberFieldFormatSwitch(fieldCode);
            if (format == "roman") {
                return PdfCore.PdfPageNumberStyle.LowerRoman;
            }

            if (format == "Roman") {
                return PdfCore.PdfPageNumberStyle.UpperRoman;
            }

            if (format == "Alphabetical") {
                return PdfCore.PdfPageNumberStyle.LowerLetter;
            }

            if (format == "ALPHABETICAL") {
                return PdfCore.PdfPageNumberStyle.UpperLetter;
            }

            if (format == "Arabic") {
                return PdfCore.PdfPageNumberStyle.Arabic;
            }

            return null;
        }

        private static string? GetNativePageNumberFieldFormatSwitch(string fieldCode) {
            int markerIndex = fieldCode.IndexOf(@"\*", StringComparison.Ordinal);
            while (markerIndex >= 0) {
                int index = markerIndex + 2;
                while (index < fieldCode.Length && char.IsWhiteSpace(fieldCode[index])) {
                    index++;
                }

                int start = index;
                while (index < fieldCode.Length && (char.IsLetter(fieldCode[index]) || fieldCode[index] == '_')) {
                    index++;
                }

                if (index > start) {
                    return fieldCode.Substring(start, index - start);
                }

                markerIndex = fieldCode.IndexOf(@"\*", markerIndex + 2, StringComparison.Ordinal);
            }

            return null;
        }

        private static void AddNativeHeaderFooterTableText(NativeHeaderFooterText parts, WordTable table, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeFontMap? nativeFontMap) {
            foreach (WordTableRow row in table.Rows) {
                IReadOnlyList<WordTableCell> cells = row.Cells;
                if (cells.Count == 1) {
                    AddNativeHeaderFooterTableCellText(parts, cells[0], NativeHeaderFooterZone.Left, listMarkers, nativeFontMap);
                    continue;
                }

                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    NativeHeaderFooterZone zone = cellIndex == 0
                        ? NativeHeaderFooterZone.Left
                        : cellIndex == cells.Count - 1
                            ? NativeHeaderFooterZone.Right
                            : NativeHeaderFooterZone.Center;

                    AddNativeHeaderFooterTableCellText(parts, cells[cellIndex], zone, listMarkers, nativeFontMap);
                }
            }
        }

        private static void AddNativeHeaderFooterTableCellText(NativeHeaderFooterText parts, WordTableCell cell, NativeHeaderFooterZone zone, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeFontMap? nativeFontMap) {
            List<WordParagraph> paragraphs = GetNativeCellParagraphs(cell).ToList();
            int lastContentIndex = -1;
            for (int index = 0; index < paragraphs.Count; index++) {
                string? text = GetNativeHeaderFooterParagraphText(paragraphs[index], listMarkers, out _);
                if (!string.IsNullOrWhiteSpace(text)) lastContentIndex = index;
            }

            for (int index = 0; index <= lastContentIndex; index++) {
                AddNativeHeaderFooterParagraphText(parts, paragraphs[index], listMarkers, zone, nativeFontMap);
            }
        }

        private static void MergeNativeHeaderFooterPageNumberStyle(ref PdfCore.PdfPageNumberStyle? current, ref bool hasConflict, PdfCore.PdfPageNumberStyle? candidate) {
            if (!candidate.HasValue || hasConflict) {
                return;
            }

            if (current.HasValue && current.Value != candidate.Value) {
                current = null;
                hasConflict = true;
                return;
            }

            current = candidate.Value;
        }

        private enum NativeHeaderFooterZone {
            Left,
            Center,
            Right
        }

        private sealed class NativeHeaderFooterFieldState {
            public bool CollectingFieldCode { get; set; }
            public bool SkippingFieldResult { get; set; }
            public StringBuilder FieldCode { get; } = new StringBuilder();
        }

        private sealed class NativeHeaderFooterImage {
            public NativeHeaderFooterImage(byte[] data, double width, double height, PdfCore.PdfAlign align) {
                Data = data;
                Width = width;
                Height = height;
                Align = align;
            }

            public byte[] Data { get; }
            public double Width { get; }
            public double Height { get; }
            public PdfCore.PdfAlign Align { get; }
        }

        private sealed class NativeHeaderFooterShape {
            public NativeHeaderFooterShape(OfficeShape shape, PdfCore.PdfAlign align) {
                Shape = shape.Clone();
                Align = align;
            }

            public OfficeShape Shape { get; }
            public PdfCore.PdfAlign Align { get; }
        }

        private sealed class NativeHeaderFooterStyledReplacement {
            public NativeHeaderFooterStyledReplacement(string serializedText, PdfCore.PdfTextRun styledRun) {
                SerializedText = serializedText;
                StyledRun = styledRun;
            }

            public string SerializedText { get; }
            public PdfCore.PdfTextRun StyledRun { get; }
        }

        private sealed class NativeHeaderFooterText {
            public string? Left { get; private set; }
            public string? Center { get; private set; }
            public string? Right { get; private set; }
            public bool HasPageTokens { get; private set; }
            public PdfCore.PdfPageNumberStyle? PageNumberStyle { get; private set; }
            public List<PdfCore.FooterSegment> LeftSegments { get; } = new List<PdfCore.FooterSegment>();
            public List<PdfCore.FooterSegment> CenterSegments { get; } = new List<PdfCore.FooterSegment>();
            public List<PdfCore.FooterSegment> RightSegments { get; } = new List<PdfCore.FooterSegment>();
            public bool HasStyledZones { get; private set; }
            private bool _hasConflictingPageNumberStyles;
            private int _leftParagraphCount;
            private int _centerParagraphCount;
            private int _rightParagraphCount;
            private bool _leftPreviousHasText;
            private bool _centerPreviousHasText;
            private bool _rightPreviousHasText;
            public bool HasContent =>
                !string.IsNullOrWhiteSpace(Left) ||
                !string.IsNullOrWhiteSpace(Center) ||
                !string.IsNullOrWhiteSpace(Right);

            public Action<PdfCore.HeaderTextBuilder>? CreateHeaderZone(NativeHeaderFooterZone zone) {
                IReadOnlyList<PdfCore.FooterSegment> segments = GetSegments(zone);
                return segments.Count == 0 ? null : builder => AppendSegments(builder, segments);
            }

            public Action<PdfCore.FooterTextBuilder>? CreateFooterZone(NativeHeaderFooterZone zone) {
                IReadOnlyList<PdfCore.FooterSegment> segments = GetSegments(zone);
                return segments.Count == 0 ? null : builder => AppendSegments(builder, segments);
            }

            public void AppendLeft(string text) => Left = Append(Left, text, null, LeftSegments, null, null, null, ref _leftParagraphCount, ref _leftPreviousHasText);
            public void AppendCenter(string text) => Center = Append(Center, text, null, CenterSegments, null, null, null, ref _centerParagraphCount, ref _centerPreviousHasText);
            public void AppendRight(string text) => Right = Append(Right, text, null, RightSegments, null, null, null, ref _rightParagraphCount, ref _rightPreviousHasText);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Left = Append(Left, text, pageNumberStyle, LeftSegments, null, null, null, ref _leftParagraphCount, ref _leftPreviousHasText);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Center = Append(Center, text, pageNumberStyle, CenterSegments, null, null, null, ref _centerParagraphCount, ref _centerPreviousHasText);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Right = Append(Right, text, pageNumberStyle, RightSegments, null, null, null, ref _rightParagraphCount, ref _rightPreviousHasText);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun) => AppendLeft(text, pageNumberStyle, markerRun, null);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun) => AppendCenter(text, pageNumberStyle, markerRun, null);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun) => AppendRight(text, pageNumberStyle, markerRun, null);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun) => AppendLeft(text, pageNumberStyle, markerRun, contentStyleRun, null);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun) => AppendCenter(text, pageNumberStyle, markerRun, contentStyleRun, null);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun) => AppendRight(text, pageNumberStyle, markerRun, contentStyleRun, null);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) => Left = Append(Left, text, pageNumberStyle, LeftSegments, markerRun, contentStyleRun, replacements, ref _leftParagraphCount, ref _leftPreviousHasText);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) => Center = Append(Center, text, pageNumberStyle, CenterSegments, markerRun, contentStyleRun, replacements, ref _centerParagraphCount, ref _centerPreviousHasText);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) => Right = Append(Right, text, pageNumberStyle, RightSegments, markerRun, contentStyleRun, replacements, ref _rightParagraphCount, ref _rightPreviousHasText);

            public void Append(NativeHeaderFooterZone zone, string text) => Append(zone, text, null);

            public void Append(NativeHeaderFooterZone zone, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) {
                switch (zone) {
                    case NativeHeaderFooterZone.Center:
                        AppendCenter(text, pageNumberStyle);
                        break;
                    case NativeHeaderFooterZone.Right:
                        AppendRight(text, pageNumberStyle);
                        break;
                    default:
                        AppendLeft(text, pageNumberStyle);
                        break;
                }
            }

            public void Append(NativeHeaderFooterZone zone, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun) {
                Append(zone, text, pageNumberStyle, markerRun, null);
            }

            public void Append(NativeHeaderFooterZone zone, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun) {
                Append(zone, text, pageNumberStyle, markerRun, contentStyleRun, null);
            }

            public void Append(NativeHeaderFooterZone zone, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) {
                switch (zone) {
                    case NativeHeaderFooterZone.Center:
                        AppendCenter(text, pageNumberStyle, markerRun, contentStyleRun, replacements);
                        break;
                    case NativeHeaderFooterZone.Right:
                        AppendRight(text, pageNumberStyle, markerRun, contentStyleRun, replacements);
                        break;
                    default:
                        AppendLeft(text, pageNumberStyle, markerRun, contentStyleRun, replacements);
                        break;
                }
            }

            public NativeHeaderFooterText Clone() {
                NativeHeaderFooterText clone = new NativeHeaderFooterText {
                    Left = Left,
                    Center = Center,
                    Right = Right,
                    HasPageTokens = HasPageTokens,
                    PageNumberStyle = PageNumberStyle,
                    _hasConflictingPageNumberStyles = _hasConflictingPageNumberStyles,
                    HasStyledZones = HasStyledZones,
                    _leftParagraphCount = _leftParagraphCount,
                    _centerParagraphCount = _centerParagraphCount,
                    _rightParagraphCount = _rightParagraphCount,
                    _leftPreviousHasText = _leftPreviousHasText,
                    _centerPreviousHasText = _centerPreviousHasText,
                    _rightPreviousHasText = _rightPreviousHasText
                };
                clone.LeftSegments.AddRange(LeftSegments);
                clone.CenterSegments.AddRange(CenterSegments);
                clone.RightSegments.AddRange(RightSegments);
                return clone;
            }

            private string Append(string? current, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle, List<PdfCore.FooterSegment> segments, PdfCore.PdfTextRun? markerRun, PdfCore.PdfTextRun? contentStyleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements, ref int paragraphCount, ref bool previousHasText) {
                text = NormalizeNativeHeaderFooterText(text);
                string plainText = (markerRun?.Text ?? string.Empty) + text;
                bool currentHasText = !string.IsNullOrWhiteSpace(plainText);
                if (text.IndexOf("{page}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("{pages}", StringComparison.OrdinalIgnoreCase) >= 0) {
                    HasPageTokens = true;
                }

                RecordPageNumberStyle(pageNumberStyle);
                string separator = previousHasText && currentHasText
                    ? Environment.NewLine + Environment.NewLine
                    : Environment.NewLine;
                if (paragraphCount > 0) {
                    segments.Add(new PdfCore.FooterSegment(PdfCore.FooterSegmentKind.Text, separator));
                }
                if (markerRun != null) {
                    segments.Add(PdfCore.FooterSegment.RichText(markerRun));
                    HasStyledZones = true;
                }
                AppendNativeHeaderFooterSegments(segments, text, contentStyleRun, replacements);
                if (contentStyleRun != null || replacements?.Count > 0) {
                    HasStyledZones = true;
                }
                string combined = paragraphCount == 0 ? plainText : (current ?? string.Empty) + separator + plainText;
                paragraphCount++;
                previousHasText = currentHasText;
                return combined;
            }

            private static void AppendNativeHeaderFooterSegments(List<PdfCore.FooterSegment> segments, string text, PdfCore.PdfTextRun? styleRun, IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements) {
                int index = 0;
                if (replacements != null) {
                    foreach (NativeHeaderFooterStyledReplacement replacement in replacements) {
                        int replacementIndex = text.IndexOf(replacement.SerializedText, index, StringComparison.Ordinal);
                        if (replacementIndex < 0) {
                            continue;
                        }

                        AppendTokenizedText(text.Substring(index, replacementIndex - index));
                        bool beginsVisualLine = replacementIndex == 0 || text[replacementIndex - 1] == '\r' || text[replacementIndex - 1] == '\n';
                        PdfCore.PdfTextRun styledRun = beginsVisualLine
                            ? replacement.StyledRun
                            : replacement.StyledRun.WithHorizontalOffset(0D);
                        segments.Add(PdfCore.FooterSegment.RichText(styledRun));
                        index = replacementIndex + replacement.SerializedText.Length;
                    }
                }
                AppendTokenizedText(text.Substring(index));

                void AppendTokenizedText(string value) {
                    int tokenCursor = 0;
                    while (tokenCursor < value.Length) {
                        int pageIndex = value.IndexOf("{page}", tokenCursor, StringComparison.OrdinalIgnoreCase);
                        int pagesIndex = value.IndexOf("{pages}", tokenCursor, StringComparison.OrdinalIgnoreCase);
                        int tokenIndex = pageIndex < 0 ? pagesIndex : pagesIndex < 0 ? pageIndex : Math.Min(pageIndex, pagesIndex);
                        if (tokenIndex < 0) {
                            if (tokenCursor < value.Length) AddText(value.Substring(tokenCursor));
                            break;
                        }
                        if (tokenIndex > tokenCursor) AddText(value.Substring(tokenCursor, tokenIndex - tokenCursor));
                        bool totalPages = tokenIndex == pagesIndex;
                        if (styleRun != null) {
                            PdfCore.PdfTextRun tokenStyle = CloneNativeHeaderFooterTextRun(styleRun, string.Empty);
                            segments.Add(totalPages ? PdfCore.FooterSegment.TotalPages(tokenStyle) : PdfCore.FooterSegment.PageNumber(tokenStyle));
                        } else {
                            segments.Add(new PdfCore.FooterSegment(totalPages ? PdfCore.FooterSegmentKind.TotalPages : PdfCore.FooterSegmentKind.PageNumber));
                        }
                        tokenCursor = tokenIndex + (totalPages ? 7 : 6);
                    }
                }

                void AddText(string value) {
                    if (styleRun != null) {
                        segments.Add(PdfCore.FooterSegment.RichText(CloneNativeHeaderFooterTextRun(styleRun, value)));
                    } else {
                        segments.Add(new PdfCore.FooterSegment(PdfCore.FooterSegmentKind.Text, value));
                    }
                }
            }

            private IReadOnlyList<PdfCore.FooterSegment> GetSegments(NativeHeaderFooterZone zone) => zone switch {
                NativeHeaderFooterZone.Center => CenterSegments,
                NativeHeaderFooterZone.Right => RightSegments,
                _ => LeftSegments
            };

            private static void AppendSegments(PdfCore.HeaderTextBuilder builder, IReadOnlyList<PdfCore.FooterSegment> segments) {
                foreach (PdfCore.FooterSegment segment in segments) {
                    if (segment.Kind == PdfCore.FooterSegmentKind.PageNumber) {
                        if (segment.StyledRun != null) builder.CurrentPage(segment.StyledRun); else builder.CurrentPage();
                    } else if (segment.Kind == PdfCore.FooterSegmentKind.TotalPages) {
                        if (segment.StyledRun != null) builder.TotalPages(segment.StyledRun); else builder.TotalPages();
                    } else if (segment.StyledRun != null) builder.Run(segment.StyledRun); else builder.Text(segment.Text ?? string.Empty);
                }
            }

            private static void AppendSegments(PdfCore.FooterTextBuilder builder, IReadOnlyList<PdfCore.FooterSegment> segments) {
                foreach (PdfCore.FooterSegment segment in segments) {
                    if (segment.Kind == PdfCore.FooterSegmentKind.PageNumber) {
                        if (segment.StyledRun != null) builder.CurrentPage(segment.StyledRun); else builder.CurrentPage();
                    } else if (segment.Kind == PdfCore.FooterSegmentKind.TotalPages) {
                        if (segment.StyledRun != null) builder.TotalPages(segment.StyledRun); else builder.TotalPages();
                    } else if (segment.StyledRun != null) builder.Run(segment.StyledRun); else builder.Text(segment.Text ?? string.Empty);
                }
            }

            private static string NormalizeNativeHeaderFooterText(string? text) {
                if (string.IsNullOrEmpty(text)) {
                    return string.Empty;
                }

                string normalized = text!
                    .Replace("\r\n", "\n")
                    .Replace('\r', '\n');
                string[] lines = normalized.Split('\n');
                for (int i = 0; i < lines.Length; i++) {
                    lines[i] = NormalizeNativeDirectText(lines[i]);
                }

                return string.Join(Environment.NewLine, lines);
            }

            private void RecordPageNumberStyle(PdfCore.PdfPageNumberStyle? style) {
                if (!style.HasValue || _hasConflictingPageNumberStyles) {
                    return;
                }

                if (PageNumberStyle.HasValue && PageNumberStyle.Value != style.Value) {
                    PageNumberStyle = null;
                    _hasConflictingPageNumberStyles = true;
                    return;
                }

                PageNumberStyle = style.Value;
            }
        }

    }
}
