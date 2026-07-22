using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const int MaximumNativeStructuredDocumentTagDepth = 128;
        private const int MaximumNativeTableNestingDepth = 128;

        [ThreadStatic]
        private static int _nativeStructuredDocumentTagDepth;

        private static void RenderNativeCoverPage(INativePdfFlow pdf, WordCoverPage coverPage, WordSection activeSection, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, double? contentWidth, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
            bool renderedCanvas = TryRenderNativeCoverPageCanvas(pdf, coverPage.Document, coverPage.SdtBlock, activeSection, options);
            RenderNativeStructuredBlockContent(pdf, coverPage.Document, coverPage.SdtBlock, activeSection, getMarker, footnoteNumbersById, options, tableOfContentsEntries, headingDestinations, contentWidth, skipCanvasOnlyVmlParagraphs: renderedCanvas, nativeDefaults, nativeFontMap);

            if (HasNativeStructuredBlockContentAfter(coverPage.SdtBlock)) {
                pdf.PageBreak();
            }
        }

        private static void RenderNativeStructuredDocumentTag(INativePdfFlow pdf, WordStructuredDocumentTag structuredDocumentTag, WordSection activeSection, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, double? contentWidth, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
            if (_nativeStructuredDocumentTagDepth >= MaximumNativeStructuredDocumentTagDepth) {
                throw new InvalidDataException(
                    $"Structured document tag nesting exceeds the supported limit of {MaximumNativeStructuredDocumentTagDepth} levels.");
            }

            _nativeStructuredDocumentTagDepth++;
            try {
                RenderNativeStructuredBlockContent(pdf, structuredDocumentTag.Document, structuredDocumentTag.SdtBlock, activeSection, getMarker, footnoteNumbersById, options, tableOfContentsEntries, headingDestinations, contentWidth, skipCanvasOnlyVmlParagraphs: false, nativeDefaults, nativeFontMap);
            } finally {
                _nativeStructuredDocumentTagDepth--;
            }
        }

        private static void RenderNativeStructuredBlockContent(INativePdfFlow pdf, WordDocument document, W.SdtBlock? sdtBlock, WordSection activeSection, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, double? contentWidth, bool skipCanvasOnlyVmlParagraphs, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
            if (TryGetNativeStructuredBlockPropertyValue(document, sdtBlock, out string? propertyValue)) {
                pdf.Paragraph(builder => builder.Text(propertyValue!));
                return;
            }

            IReadOnlyList<WordElement> elements = CollapseNativeParagraphElements(GetNativeStructuredBlockElements(document, sdtBlock, skipCanvasOnlyVmlParagraphs));
            for (int i = 0; i < elements.Count; i++) {
                WordElement element = elements[i];
                if (element is WordFootNote) {
                    continue;
                }

                RenderNativeElement(
                    pdf,
                    element,
                    activeSection,
                    getMarker,
                    GetNativeStructuredBlockFootnoteNumbersForElement(elements, i, footnoteNumbersById),
                    footnoteNumbersById,
                    options,
                    tableOfContentsEntries,
                    headingDestinations,
                    contentWidth,
                    nativeDefaults,
                    nativeFontMap,
                    nextElement: GetNextNativeRenderableElement(elements, i));
            }
        }

        private static IReadOnlyList<WordElement> GetNativeStructuredBlockElements(WordDocument document, W.SdtBlock? sdtBlock, bool skipCanvasOnlyVmlParagraphs = false) {
            var elements = new List<WordElement>();
            W.SdtContentBlock? content = sdtBlock?.SdtContentBlock;
            if (content == null) {
                return elements;
            }

            foreach (OpenXmlElement child in content.ChildElements) {
                if (child is W.Paragraph paragraph) {
                    if (skipCanvasOnlyVmlParagraphs && IsNativeCanvasOnlyVmlParagraph(paragraph)) {
                        continue;
                    }

                    elements.AddRange(WordSection.ConvertParagraphToWordParagraphs(document, paragraph));
                } else if (child is W.Table table) {
                    elements.Add(new WordTable(document, table));
                } else if (child is W.SdtBlock nestedSdtBlock) {
                    elements.Add(WordSection.ConvertStdBlockToWordElements(document, nestedSdtBlock));
                }
            }

            return elements;
        }

        private static IEnumerable<WordTable> EnumerateNativeTableTree(WordTable root) {
            var pending = new Stack<(WordTable Table, int Depth)>();
            pending.Push((root, 0));
            while (pending.Count > 0) {
                (WordTable table, int depth) = pending.Pop();
                EnsureNativeTableDepth(depth);
                yield return table;

                List<WordTableRow> rows = table.Rows;
                for (int rowIndex = rows.Count - 1; rowIndex >= 0; rowIndex--) {
                    List<WordTableCell> cells = rows[rowIndex].Cells;
                    for (int cellIndex = cells.Count - 1; cellIndex >= 0; cellIndex--) {
                        List<WordTable> nestedTables = cells[cellIndex].DirectNestedTables;
                        for (int nestedIndex = nestedTables.Count - 1; nestedIndex >= 0; nestedIndex--) {
                            pending.Push((nestedTables[nestedIndex], depth + 1));
                        }
                    }
                }
            }
        }

        private static void EnsureNativeTableDepth(int depth) {
            if (depth >= MaximumNativeTableNestingDepth) {
                throw new InvalidDataException(
                    $"Table nesting exceeds the supported limit of {MaximumNativeTableNestingDepth} levels.");
            }
        }

        private static IReadOnlyList<int> GetNativeStructuredBlockFootnoteNumbersForElement(IReadOnlyList<WordElement> elements, int index, Dictionary<long, int> footnoteNumbersById) {
            var numbers = GetNativeFootnoteNumbersForElement(elements, index, footnoteNumbersById).ToList();
            var leadingNumbers = new List<int>();
            int previousIndex = index - 1;
            while (previousIndex >= 0 && (elements[previousIndex] is WordFootNote || elements[previousIndex] is WordEndNote)) {
                long? key = GetNativeNoteKey(elements[previousIndex]);
                if (key.HasValue && footnoteNumbersById.TryGetValue(key.Value, out int number)) {
                    leadingNumbers.Add(number);
                }

                previousIndex--;
            }

            if (previousIndex < 0) {
                numbers.AddRange(leadingNumbers);
            }

            return numbers.Distinct().ToList();
        }

        private static bool IsNativeCanvasOnlyVmlParagraph(W.Paragraph paragraph) {
            if (!HasNativeVmlCoverDrawing(paragraph)) {
                return false;
            }

            foreach (W.Text text in paragraph.Descendants<W.Text>()) {
                if (string.IsNullOrWhiteSpace(text.Text)) {
                    continue;
                }

                bool insideTextBox = text.Ancestors().Any(ancestor =>
                    ancestor.NamespaceUri == "urn:schemas-microsoft-com:vml" ||
                    ancestor.LocalName == "txbxContent");
                if (!insideTextBox) {
                    return false;
                }
            }

            return true;
        }

        private static bool TryGetNativeStructuredBlockPropertyValue(WordDocument document, W.SdtBlock? sdtBlock, out string? value) {
            value = null;
            W.SdtProperties? properties = sdtBlock?.SdtProperties;
            if (properties == null) {
                return false;
            }

            if (!IsNativePropertyBoundStructuredBlock(properties)) {
                return false;
            }

            value = GetNativeBuiltInPropertyValue(document, properties);
            return !string.IsNullOrWhiteSpace(value);
        }

        private static bool IsNativePropertyBoundStructuredBlock(W.SdtProperties properties) {
            W.DataBinding? binding = properties.Elements<W.DataBinding>().FirstOrDefault();
            if (binding != null && !string.IsNullOrWhiteSpace(binding.XPath?.Value)) {
                return true;
            }

            return properties.Elements<W.ShowingPlaceholder>().Any();
        }

        private static bool HasNativeStructuredBlockContentAfter(W.SdtBlock? sdtBlock) {
            if (sdtBlock == null) {
                return false;
            }

            for (OpenXmlElement? next = sdtBlock.NextSibling(); next != null; next = next.NextSibling()) {
                if (next is W.SectionProperties) {
                    continue;
                }

                if (next is W.Paragraph paragraph) {
                    if (IsNativePageBreakSeparator(paragraph)) {
                        return false;
                    }

                    if (IsNativeEmptyParagraph(paragraph)) {
                        continue;
                    }
                }

                return true;
            }

            return false;
        }

        private static bool IsNativeEmptyParagraph(W.Paragraph paragraph) {
            if (paragraph.Descendants<W.Drawing>().Any() ||
                paragraph.Descendants().Any(element => element.LocalName == "pict" || element.NamespaceUri == "urn:schemas-microsoft-com:vml")) {
                return false;
            }

            return !paragraph.Descendants<W.Text>().Any(text => !string.IsNullOrWhiteSpace(text.Text));
        }

        private static bool IsNativePageBreakSeparator(W.Paragraph paragraph) =>
            paragraph.ParagraphProperties?.PageBreakBefore != null ||
            paragraph.Descendants<W.Break>().Any(breakElement => breakElement.Type?.Value == W.BreakValues.Page);

        private static string? GetNativeBuiltInPropertyValue(WordDocument document, W.SdtProperties? properties) {
            if (properties == null) {
                return null;
            }

            string? alias = properties.Elements<W.SdtAlias>().FirstOrDefault()?.Val?.Value;
            string? xPath = properties.Elements<W.DataBinding>().FirstOrDefault()?.XPath?.Value;
            return GetNativeBuiltInPropertyValue(document, alias, xPath);
        }

        private static string? GetNativeBuiltInPropertyValue(WordDocument document, string? alias, string? xPath) {
            string key = (alias ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(xPath)) {
                key = xPath!;
            }

            if (key.IndexOf("Company", StringComparison.OrdinalIgnoreCase) >= 0) {
                return EmptyToNull(document.ApplicationProperties?.Company);
            }

            if (key.Equals("Title", StringComparison.OrdinalIgnoreCase) ||
                key.IndexOf("Title[", StringComparison.OrdinalIgnoreCase) >= 0) {
                return EmptyToNull(document.BuiltinDocumentProperties?.Title);
            }

            if (key.Equals("Subtitle", StringComparison.OrdinalIgnoreCase) ||
                key.Equals("Subject", StringComparison.OrdinalIgnoreCase)) {
                return EmptyToNull(document.BuiltinDocumentProperties?.Subject);
            }

            if (key.Equals("Author", StringComparison.OrdinalIgnoreCase) ||
                key.Equals("Creator", StringComparison.OrdinalIgnoreCase)) {
                return EmptyToNull(document.BuiltinDocumentProperties?.Creator);
            }

            if (key.Equals("Date", StringComparison.OrdinalIgnoreCase)) {
                DateTime? date = document.BuiltinDocumentProperties?.Created ??
                                 document.BuiltinDocumentProperties?.Modified;
                return date?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            return null;
        }

        private static string? EmptyToNull(string? value) =>
            string.IsNullOrWhiteSpace(value) ? null : value;

        private static string ResolveNativeBuiltInPropertyPlaceholders(WordDocument document, string text) {
            if (string.IsNullOrEmpty(text)) {
                return text;
            }

            text = ReplaceNativeBuiltInPropertyPlaceholder(document, text, "[company name]", "Company", null);
            text = ReplaceNativeBuiltInPropertyPlaceholder(document, text, "[Company name]", "Company", null);
            text = ReplaceNativeBuiltInPropertyPlaceholder(document, text, "[Company Name]", "Company", null);
            text = ReplaceNativeBuiltInPropertyPlaceholder(document, text, "[Document title]", "Title", null);
            text = ReplaceNativeBuiltInPropertyPlaceholder(document, text, "[Document Title]", "Title", null);
            text = ReplaceNativeBuiltInPropertyPlaceholder(document, text, "[Document subtitle]", "Subtitle", null);
            text = ReplaceNativeBuiltInPropertyPlaceholder(document, text, "[Document Subtitle]", "Subtitle", null);
            text = ReplaceNativeBuiltInPropertyPlaceholder(document, text, "[Author]", "Author", null);
            text = ReplaceNativeBuiltInPropertyPlaceholder(document, text, "[Date]", "Date", null);
            return text;
        }

        private static string ReplaceNativeBuiltInPropertyPlaceholder(WordDocument document, string text, string placeholder, string alias, string? xPath) {
            if (text.IndexOf(placeholder, StringComparison.Ordinal) < 0) {
                return text;
            }

            string? value = GetNativeBuiltInPropertyValue(document, alias, xPath);
            return string.IsNullOrWhiteSpace(value)
                ? text
                : text.Replace(placeholder, value);
        }

        private static void ApplyNativeSectionWatermark(PdfCore.PdfPageCompose page, WordSection section, PdfSaveOptions? options) {
            if (HasNativeHeaderSpecificWatermarks(section)) {
                WordWatermark? defaultWatermark = section.Header?.Default?.Watermarks.FirstOrDefault(IsNativeRenderableWatermark);
                NativeAppliedWatermark defaultApplied = ApplyNativeWatermark(page.Watermark, page.ImageWatermark, defaultWatermark, options, "default header watermark");
                if (section.DifferentFirstPage) {
                    WordWatermark? firstWatermark = section.Header?.First?.Watermarks.FirstOrDefault(IsNativeRenderableWatermark);
                    NativeAppliedWatermark firstApplied = ApplyNativeWatermark(page.FirstPageWatermark, page.FirstPageImageWatermark, firstWatermark, options, "first header watermark");
                    SuppressMissingFirstPageWatermark(page, defaultApplied, firstApplied);
                }

                if (section.DifferentOddAndEvenPages) {
                    WordWatermark? evenWatermark = section.Header?.Even?.Watermarks.FirstOrDefault(IsNativeRenderableWatermark);
                    NativeAppliedWatermark evenApplied = ApplyNativeWatermark(page.EvenPagesWatermark, page.EvenPagesImageWatermark, evenWatermark, options, "even header watermark");
                    SuppressMissingEvenPagesWatermark(page, defaultApplied, evenApplied);
                }

                return;
            }

            WordWatermark? watermark = section.Header?.Default?.Watermarks.FirstOrDefault(IsNativeRenderableWatermark);
            _ = ApplyNativeWatermark(page.Watermark, page.ImageWatermark, watermark, options, "default header watermark");
        }

        private static NativeAppliedWatermark ApplyNativeWatermark(
            Func<PdfCore.PdfTextWatermark?, PdfCore.PdfPageCompose> applyText,
            Func<PdfCore.PdfImageWatermark?, PdfCore.PdfPageCompose> applyImage,
            WordWatermark? watermark,
            PdfSaveOptions? options,
            string source) {
            if (watermark == null) {
                return default;
            }

            PdfCore.PdfTextWatermark? textWatermark = CreateNativeTextWatermark(watermark);
            PdfCore.PdfImageWatermark? imageWatermark = CreateNativeImageWatermark(watermark, options, source);
            applyText(textWatermark);
            applyImage(imageWatermark);
            return new NativeAppliedWatermark(textWatermark != null, imageWatermark != null);
        }

        private static void SuppressMissingFirstPageWatermark(PdfCore.PdfPageCompose page, NativeAppliedWatermark defaultWatermark, NativeAppliedWatermark firstWatermark) {
            if (defaultWatermark.HasText && !firstWatermark.HasText) {
                page.SuppressFirstPageTextWatermark();
            }

            if (defaultWatermark.HasImage && !firstWatermark.HasImage) {
                page.SuppressFirstPageImageWatermark();
            }
        }

        private static void SuppressMissingEvenPagesWatermark(PdfCore.PdfPageCompose page, NativeAppliedWatermark defaultWatermark, NativeAppliedWatermark evenWatermark) {
            if (defaultWatermark.HasText && !evenWatermark.HasText) {
                page.SuppressEvenPagesTextWatermark();
            }

            if (defaultWatermark.HasImage && !evenWatermark.HasImage) {
                page.SuppressEvenPagesImageWatermark();
            }
        }

        private readonly struct NativeAppliedWatermark {
            public NativeAppliedWatermark(bool hasText, bool hasImage) {
                HasText = hasText;
                HasImage = hasImage;
            }

            public bool HasText { get; }
            public bool HasImage { get; }
        }

        private static PdfCore.PdfTextWatermark? CreateNativeTextWatermark(WordWatermark? watermark) {
            if (watermark == null) {
                return null;
            }

            if (string.IsNullOrWhiteSpace(watermark.Text)) {
                return null;
            }

            var pdfWatermark = new PdfCore.PdfTextWatermark(watermark.Text) {
                Bold = true
            };
            if (watermark.Rotation.HasValue) {
                pdfWatermark.RotationAngle = watermark.Rotation.Value;
            }

            PdfCore.PdfColor? color = ParseNativeWatermarkColor(watermark.ColorHex);
            if (color.HasValue) {
                pdfWatermark.Color = color.Value;
            }

            double? opacity = watermark.Opacity;
            if (opacity.HasValue) {
                pdfWatermark.Opacity = opacity.Value;
            }

            double? fontSize = ResolveNativeWatermarkFontSize(watermark);
            if (fontSize.HasValue) {
                pdfWatermark.FontSize = fontSize.Value;
            }

            if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(watermark.FontFamily, out PdfCore.PdfStandardFont mappedFont)) {
                pdfWatermark.Font = mappedFont;
            }

            return pdfWatermark;
        }

        private static PdfCore.PdfImageWatermark? CreateNativeImageWatermark(WordWatermark? watermark, PdfSaveOptions? options, string source) {
            if (watermark == null || !watermark.HasImage) {
                return null;
            }

            if (!watermark.TryGetImageBytes(out byte[] bytes, out string? unsupportedReason)) {
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "NativeWatermarkImageUnsupported",
                        source,
                        "Word image watermark was not exported because the image part could not be read. " + unsupportedReason);
                }

                return null;
            }

            if (!TryPrepareNativePdfImageBytes(bytes, out byte[] preparedBytes, out unsupportedReason)) {
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "NativeWatermarkImageUnsupported",
                        source,
                        "Word image watermark was not exported because the shared PDF raster pipeline could not prepare it. " + unsupportedReason);
                }

                return null;
            }

            double width = watermark.Width is > 0D ? watermark.Width.Value : 144D;
            double height = watermark.Height is > 0D ? watermark.Height.Value : 144D;
            var pdfWatermark = new PdfCore.PdfImageWatermark(preparedBytes, width, height);
            if (watermark.Rotation.HasValue) {
                pdfWatermark.RotationAngle = watermark.Rotation.Value;
            }

            double? opacity = watermark.Opacity;
            if (opacity.HasValue) {
                pdfWatermark.Opacity = opacity.Value;
            }

            return pdfWatermark;
        }

        private static bool HasNativeHeaderSpecificWatermarks(WordSection section) {
            if (!section.DifferentFirstPage && !section.DifferentOddAndEvenPages) {
                return false;
            }

            return HasNativeWatermark(section.Header?.Default) ||
                   HasNativeWatermark(section.Header?.First) ||
                   HasNativeWatermark(section.Header?.Even);
        }

        private static bool HasNativeWatermark(WordHeaderFooter? headerFooter) =>
            headerFooter?.Watermarks.Any(IsNativeRenderableWatermark) == true;

        private static bool IsNativeRenderableWatermark(WordWatermark mark) =>
            !string.IsNullOrWhiteSpace(mark.Text) || mark.HasImage;

        private static double? ResolveNativeWatermarkFontSize(WordWatermark watermark) {
            double? fontSize = watermark.FontSize;
            if (fontSize.HasValue && fontSize.Value > 2D) {
                return fontSize.Value;
            }

            double? width = watermark.Width;
            double? height = watermark.Height;
            string text = watermark.Text ?? string.Empty;
            if (!width.HasValue || !height.HasValue || width.Value <= 0D || height.Value <= 0D || text.Length == 0) {
                return fontSize;
            }

            double widthBound = width.Value / Math.Max(1D, text.Length * 0.58D);
            double heightBound = height.Value * 0.72D;
            double derived = Math.Min(widthBound, heightBound);
            return derived > 2D ? derived : fontSize;
        }

        private static PdfCore.PdfColor? ParseNativeWatermarkColor(string? color) {
            PdfCore.PdfColor? parsed = ParseNativeColor(color);
            if (parsed != null || string.IsNullOrWhiteSpace(color)) {
                return parsed;
            }

            string value = color!.Trim();
            if (value.Equals("silver", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfColor.FromRgb(192, 192, 192);
            if (value.Equals("gray", StringComparison.OrdinalIgnoreCase) || value.Equals("grey", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfColor.Gray;
            if (value.Equals("black", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfColor.Black;
            if (value.Equals("white", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfColor.White;
            if (value.Equals("red", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfColor.FromRgb(255, 0, 0);
            if (value.Equals("green", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfColor.FromRgb(0, 128, 0);
            if (value.Equals("blue", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfColor.FromRgb(0, 0, 255);

            return null;
        }
    }
}
