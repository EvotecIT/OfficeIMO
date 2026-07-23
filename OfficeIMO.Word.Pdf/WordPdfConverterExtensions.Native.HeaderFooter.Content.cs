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
        private static void ApplyNativeHeaderFooterPageNumberStyle(PdfCore.PdfPageCompose page, params NativeHeaderFooterText?[] parts) {
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

        private static NativeHeaderFooterText? GetNativeHeaderFooterText(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                return null;
            }

            var parts = new NativeHeaderFooterText();
            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                switch (element) {
                    case WordParagraph paragraph:
                        AddNativeHeaderFooterParagraphText(parts, paragraph);
                        break;
                    case WordTable table:
                        AddNativeHeaderFooterTableText(parts, table);
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
                foreach (string familyName in EnumerateNativeHeaderFooterFontFamilies(headerFooter)) {
                    if (!TryResolveNativeMappedFont(familyName, nativeFontMap, out PdfCore.PdfStandardFont mappedFont)) {
                        continue;
                    }

                    PdfCore.PdfStandardFont fontFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(mappedFont);
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
                foreach (string familyName in EnumerateNativeHeaderFooterFontFamilies(headerFooter)) {
                    if (!nativeFontMap.TryGetNamedFontFamily(familyName, out string? namedFamily) ||
                        string.IsNullOrWhiteSpace(namedFamily)) {
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
                foreach (PdfCore.PdfColor color in EnumerateNativeHeaderFooterColors(headerFooter)) {
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

        private static PdfCore.PdfStandardFont ResolveNativeHeaderFooterBaseFont(WordDocument document, PdfSaveOptions? options, bool isHeader) {
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

        private static IEnumerable<NativeResolvedTextStyle> EnumerateNativeHeaderFooterTextStyles(WordHeaderFooter? headerFooter) {
            if (headerFooter == null) {
                yield break;
            }

            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeHeaderFooterElementTextStyles(element)) {
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

        private static IEnumerable<NativeResolvedTextStyle> EnumerateNativeHeaderFooterElementTextStyles(WordElement element) {
            if (element is WordParagraph paragraph) {
                foreach (NativeResolvedTextStyle style in EnumerateNativeParagraphTextStyles(paragraph)) {
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
                            foreach (NativeResolvedTextStyle style in EnumerateNativeParagraphTextStyles(cellParagraph)) {
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

        private static IEnumerable<NativeResolvedTextStyle> EnumerateNativeParagraphTextStyles(WordParagraph paragraph) {
            List<WordParagraph> runs = GetNativeRuns(paragraph);
            bool emittedRun = false;
            foreach (WordParagraph run in runs) {
                if (run.IsImage || string.IsNullOrWhiteSpace(run.Text)) {
                    continue;
                }

                emittedRun = true;
                yield return ResolveNativeTextRunStyle(run, paragraph);
            }

            if (!emittedRun && !string.IsNullOrWhiteSpace(paragraph.Text)) {
                yield return ResolveNativeTextRunStyle(paragraph);
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

        private static IReadOnlyList<NativeHeaderFooterImage> GetNativeHeaderFooterImages(WordHeaderFooter? headerFooter, PdfSaveOptions? options, string source) {
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

        private static void AddNativeHeaderFooterParagraphText(NativeHeaderFooterText parts, WordParagraph paragraph) {
            string? text = GetNativeHeaderFooterParagraphText(paragraph, out PdfCore.PdfPageNumberStyle? pageNumberStyle, out NativeHeaderFooterZone? zoneOverride);

            if (string.IsNullOrWhiteSpace(text)) {
                return;
            }

            string resolvedText = text!;
            if (zoneOverride.HasValue) {
                parts.Append(zoneOverride.Value, resolvedText, pageNumberStyle);
                return;
            }

            W.JustificationValues? alignment = ResolveNativeParagraphJustification(paragraph);
            if (alignment == W.JustificationValues.Center) {
                parts.AppendCenter(resolvedText, pageNumberStyle);
            } else if (alignment == W.JustificationValues.Right) {
                parts.AppendRight(resolvedText, pageNumberStyle);
            } else {
                parts.AppendLeft(resolvedText, pageNumberStyle);
            }
        }

        private static void AddNativeHeaderFooterTableImages(List<NativeHeaderFooterImage> images, WordTable table, PdfSaveOptions? options, string source) {
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

        private static void AddNativeHeaderFooterParagraphImage(List<NativeHeaderFooterImage> images, WordParagraph paragraph, PdfCore.PdfAlign? alignOverride, PdfSaveOptions? options, string source) {
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

        private static void AddNativeHeaderFooterImage(List<NativeHeaderFooterImage> images, WordImage image, PdfCore.PdfAlign align, PdfSaveOptions? options, string source) {
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

        private static string? GetNativeHeaderFooterParagraphText(WordParagraph paragraph, out PdfCore.PdfPageNumberStyle? pageNumberStyle) {
            return GetNativeHeaderFooterParagraphText(paragraph, out pageNumberStyle, out _);
        }

        private static string? GetNativeHeaderFooterParagraphText(WordParagraph paragraph, out PdfCore.PdfPageNumberStyle? pageNumberStyle, out NativeHeaderFooterZone? zoneOverride) {
            zoneOverride = null;
            if (TryBuildNativeHeaderFooterParagraphText(paragraph, out string? mixedText, out pageNumberStyle)) {
                return AppendNativeHeaderFooterSupplementalText(mixedText, paragraph);
            }

            if (TryGetNativeHeaderFooterFieldToken(paragraph, out string? fieldToken, out pageNumberStyle)) {
                return AppendNativeHeaderFooterSupplementalText(fieldToken, paragraph);
            }

            pageNumberStyle = null;
            if (paragraph.IsHyperLink && paragraph.Hyperlink != null && !IsNativeHiddenTextRun(paragraph)) {
                return AppendNativeHeaderFooterSupplementalText(ApplyNativeTextTransform(paragraph.Hyperlink.Text, paragraph), paragraph);
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            string? text = runs.Count > 0
                ? string.Concat(runs.Where(run => !IsNativeHiddenTextRun(run, paragraph)).Select(run => ApplyNativeTextTransform(run.Text, run, paragraph)))
                : IsNativeHiddenTextRun(paragraph) ? string.Empty : ApplyNativeTextTransform(paragraph.Text, paragraph);
            text = AppendNativeHeaderFooterSupplementalText(text, paragraph);
            if (!string.IsNullOrWhiteSpace(text)) {
                return text;
            }

            string? textBoxText = GetNativeParagraphTextBoxPlainText(paragraph);
            if (string.IsNullOrWhiteSpace(textBoxText)) {
                return text;
            }

            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out _);
            zoneOverride = MapNativeTextBoxHeaderFooterZone(textBox?.HorizontalAlignment ?? WordHorizontalAlignmentValues.Center);
            return textBoxText;
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

        private static NativeHeaderFooterZone MapNativeTextBoxHeaderFooterZone(WordHorizontalAlignmentValues alignment) {
            switch (alignment) {
                case WordHorizontalAlignmentValues.Center:
                    return NativeHeaderFooterZone.Center;
                case WordHorizontalAlignmentValues.Right:
                case WordHorizontalAlignmentValues.Outside:
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

        private static void AddNativeHeaderFooterTableText(NativeHeaderFooterText parts, WordTable table) {
            foreach (WordTableRow row in table.Rows) {
                IReadOnlyList<WordTableCell> cells = row.Cells;
                if (cells.Count == 1) {
                    AddNativeHeaderFooterSingleCellText(parts, cells[0]);
                    continue;
                }

                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    NativeHeaderFooterZone zone = cellIndex == 0
                        ? NativeHeaderFooterZone.Left
                        : cellIndex == cells.Count - 1
                            ? NativeHeaderFooterZone.Right
                            : NativeHeaderFooterZone.Center;

                    AddNativeHeaderFooterCellText(parts, cells[cellIndex], zone);
                }
            }
        }

        private static void AddNativeHeaderFooterSingleCellText(NativeHeaderFooterText parts, WordTableCell cell) {
            string cellText = GetNativeHeaderFooterCellText(cell, preserveParagraphBreaks: true, out PdfCore.PdfPageNumberStyle? pageNumberStyle);
            if (!string.IsNullOrWhiteSpace(cellText)) {
                parts.AppendLeft(cellText, pageNumberStyle);
            }
        }

        private static void AddNativeHeaderFooterCellText(NativeHeaderFooterText parts, WordTableCell cell, NativeHeaderFooterZone zone) {
            string cellText = GetNativeHeaderFooterCellText(cell, preserveParagraphBreaks: true, out PdfCore.PdfPageNumberStyle? pageNumberStyle);
            if (!string.IsNullOrWhiteSpace(cellText)) {
                parts.Append(zone, cellText, pageNumberStyle);
            }
        }

        private static string GetNativeHeaderFooterCellText(WordTableCell cell, bool preserveParagraphBreaks, out PdfCore.PdfPageNumberStyle? pageNumberStyle) {
            var parts = new List<string>();
            pageNumberStyle = null;
            bool hasConflictingStyles = false;
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                string? text = GetNativeHeaderFooterParagraphText(paragraph, out PdfCore.PdfPageNumberStyle? paragraphStyle);
                if (!string.IsNullOrEmpty(text)) {
                    parts.Add(text!);
                    MergeNativeHeaderFooterPageNumberStyle(ref pageNumberStyle, ref hasConflictingStyles, paragraphStyle);
                } else {
                    parts.Add(string.Empty);
                }
            }

            while (parts.Count > 0 && string.IsNullOrWhiteSpace(parts[parts.Count - 1])) {
                parts.RemoveAt(parts.Count - 1);
            }

            if (!parts.Any(part => !string.IsNullOrWhiteSpace(part))) {
                return string.Empty;
            }

            return JoinNativeHeaderFooterParagraphParts(parts, preserveParagraphBreaks);
        }

        private static string JoinNativeHeaderFooterParagraphParts(IReadOnlyList<string> parts, bool preserveParagraphBreaks) {
            var builder = new StringBuilder();
            for (int index = 0; index < parts.Count; index++) {
                string part = parts[index];
                if (index > 0) {
                    bool previousHasText = !string.IsNullOrWhiteSpace(parts[index - 1]);
                    bool currentHasText = !string.IsNullOrWhiteSpace(part);
                    if (preserveParagraphBreaks) {
                        builder.Append(Environment.NewLine);
                        if (previousHasText && currentHasText) {
                            builder.Append(Environment.NewLine);
                        }
                    } else if (previousHasText != currentHasText) {
                        builder.Append(Environment.NewLine);
                    } else if (previousHasText && currentHasText) {
                        builder.Append(' ');
                    }
                }

                builder.Append(part);
            }

            return builder.ToString();
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

        private sealed class NativeHeaderFooterText {
            public string? Left { get; private set; }
            public string? Center { get; private set; }
            public string? Right { get; private set; }
            public bool HasPageTokens { get; private set; }
            public PdfCore.PdfPageNumberStyle? PageNumberStyle { get; private set; }
            private bool _hasConflictingPageNumberStyles;
            public bool HasContent =>
                !string.IsNullOrWhiteSpace(Left) ||
                !string.IsNullOrWhiteSpace(Center) ||
                !string.IsNullOrWhiteSpace(Right);

            public void AppendLeft(string text) => Left = Append(Left, text, null);
            public void AppendCenter(string text) => Center = Append(Center, text, null);
            public void AppendRight(string text) => Right = Append(Right, text, null);
            public void AppendLeft(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Left = Append(Left, text, pageNumberStyle);
            public void AppendCenter(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Center = Append(Center, text, pageNumberStyle);
            public void AppendRight(string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) => Right = Append(Right, text, pageNumberStyle);

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

            public NativeHeaderFooterText Clone() {
                return new NativeHeaderFooterText {
                    Left = Left,
                    Center = Center,
                    Right = Right,
                    HasPageTokens = HasPageTokens,
                    PageNumberStyle = PageNumberStyle,
                    _hasConflictingPageNumberStyles = _hasConflictingPageNumberStyles
                };
            }

            private string Append(string? current, string text, PdfCore.PdfPageNumberStyle? pageNumberStyle) {
                text = NormalizeNativeHeaderFooterText(text);
                if (text.IndexOf("{page}", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("{pages}", StringComparison.OrdinalIgnoreCase) >= 0) {
                    HasPageTokens = true;
                }

                RecordPageNumberStyle(pageNumberStyle);
                if (string.IsNullOrWhiteSpace(current)) {
                    return text;
                }

                string separator = !string.IsNullOrWhiteSpace(text)
                    ? Environment.NewLine + Environment.NewLine
                    : Environment.NewLine;
                return current + separator + text;
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
