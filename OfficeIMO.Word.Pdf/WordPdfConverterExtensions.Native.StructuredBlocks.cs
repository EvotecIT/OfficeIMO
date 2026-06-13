using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void RenderNativeCoverPage(INativePdfFlow pdf, WordCoverPage coverPage, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, double? contentWidth) {
            bool renderedCanvas = TryRenderNativeCoverPageCanvas(pdf, coverPage.Document, coverPage.SdtBlock, options);
            RenderNativeStructuredBlockContent(pdf, coverPage.Document, coverPage.SdtBlock, getMarker, footnoteNumbersById, options, tableOfContentsEntries, headingDestinations, contentWidth, skipCanvasOnlyVmlParagraphs: renderedCanvas);

            pdf.PageBreak();
        }

        private static void RenderNativeStructuredDocumentTag(INativePdfFlow pdf, WordStructuredDocumentTag structuredDocumentTag, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, double? contentWidth) {
            RenderNativeStructuredBlockContent(pdf, structuredDocumentTag.Document, structuredDocumentTag.SdtBlock, getMarker, footnoteNumbersById, options, tableOfContentsEntries, headingDestinations, contentWidth, skipCanvasOnlyVmlParagraphs: false);
        }

        private static void RenderNativeStructuredBlockContent(INativePdfFlow pdf, WordDocument document, W.SdtBlock? sdtBlock, Func<WordParagraph, (int Level, string Marker)?> getMarker, Dictionary<long, int> footnoteNumbersById, PdfSaveOptions? options, IReadOnlyList<NativeTableOfContentsEntry> tableOfContentsEntries, IReadOnlyDictionary<W.Paragraph, string> headingDestinations, double? contentWidth, bool skipCanvasOnlyVmlParagraphs) {
            if (TryGetNativeStructuredBlockPropertyValue(document, sdtBlock, out string? propertyValue)) {
                pdf.Paragraph(builder => builder.Text(propertyValue!));
                return;
            }

            IReadOnlyList<WordElement> elements = CollapseNativeParagraphElements(GetNativeStructuredBlockElements(document, sdtBlock, skipCanvasOnlyVmlParagraphs));
            foreach (WordElement element in elements) {
                if (element is WordFootNote) {
                    continue;
                }

                RenderNativeElement(
                    pdf,
                    element,
                    getMarker,
                    Array.Empty<int>(),
                    footnoteNumbersById,
                    options,
                    tableOfContentsEntries,
                    headingDestinations,
                    contentWidth);
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

            value = GetNativeBuiltInPropertyValue(document, properties);
            return !string.IsNullOrWhiteSpace(value);
        }

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
                ApplyNativeWatermark(page.Watermark, page.ImageWatermark, section.Header?.Default?.Watermarks.FirstOrDefault(IsNativeRenderableWatermark), options, "default header watermark");
                if (section.DifferentFirstPage) {
                    ApplyNativeWatermark(page.FirstPageWatermark, page.FirstPageImageWatermark, section.Header?.First?.Watermarks.FirstOrDefault(IsNativeRenderableWatermark), options, "first header watermark");
                }

                if (section.DifferentOddAndEvenPages) {
                    ApplyNativeWatermark(page.EvenPagesWatermark, page.EvenPagesImageWatermark, section.Header?.Even?.Watermarks.FirstOrDefault(IsNativeRenderableWatermark), options, "even header watermark");
                }

                return;
            }

            WordWatermark? watermark = section.Watermarks.FirstOrDefault(IsNativeRenderableWatermark);
            ApplyNativeWatermark(page.Watermark, page.ImageWatermark, watermark, options, "section watermark");
        }

        private static void ApplyNativeWatermark(
            Func<PdfCore.PdfTextWatermark?, PdfCore.PdfPageCompose> applyText,
            Func<PdfCore.PdfImageWatermark?, PdfCore.PdfPageCompose> applyImage,
            WordWatermark? watermark,
            PdfSaveOptions? options,
            string source) {
            PdfCore.PdfTextWatermark? textWatermark = CreateNativeTextWatermark(watermark);
            PdfCore.PdfImageWatermark? imageWatermark = CreateNativeImageWatermark(watermark, options, source);
            applyText(textWatermark);
            applyImage(imageWatermark);
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

            if (!IsNativePdfSupportedImageBytes(bytes, out unsupportedReason)) {
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "NativeWatermarkImageUnsupported",
                        source,
                        "Word image watermark was not exported because the first-party PDF image writer supports JPEG and simple PNG images only. " + unsupportedReason);
                }

                return null;
            }

            double width = watermark.Width is > 0D ? watermark.Width.Value : 144D;
            double height = watermark.Height is > 0D ? watermark.Height.Value : 144D;
            var pdfWatermark = new PdfCore.PdfImageWatermark(bytes, width, height);
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
