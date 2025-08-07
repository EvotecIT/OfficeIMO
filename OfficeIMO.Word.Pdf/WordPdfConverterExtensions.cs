using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {

    /// <summary>
    /// Provides extension methods for converting <see cref="WordDocument"/> instances to PDF files.
    /// </summary>
    public static partial class WordPdfConverterExtensions {
        /// <summary>
        /// Saves the specified <see cref="WordDocument"/> as a PDF at the given <paramref name="path"/>.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="path">The output PDF file path.</param>
        /// <param name="options">Optional PDF configuration.</param>
        public static void SaveAsPdf(this WordDocument document, string path, PdfSaveOptions? options = null) {
            Document pdf = CreatePdfDocument(document, options);
            pdf.GeneratePdf(path);
        }

        /// <summary>
        /// Saves the specified <see cref="WordDocument"/> as a PDF to the provided <paramref name="stream"/>.
        /// </summary>
        /// <param name="document">The document to convert.</param>
        /// <param name="stream">The output stream to receive the PDF data.</param>
        /// <param name="options">Optional PDF configuration.</param>
        public static void SaveAsPdf(this WordDocument document, Stream stream, PdfSaveOptions? options = null) {
            Document pdf = CreatePdfDocument(document, options);
            pdf.GeneratePdf(stream);
        }

        private static Document CreatePdfDocument(WordDocument document, PdfSaveOptions? options) {
            QuestPDF.Settings.License = LicenseType.Community;

            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers = DocumentTraversal.BuildListMarkers(document);

            Document pdf = Document.Create(container => {
                foreach (WordSection section in document.Sections) {
                    container.Page(page => {
                        if (!string.IsNullOrEmpty(options?.FontFamily)) {
                            page.DefaultTextStyle(t => t.FontFamily(options.FontFamily));
                        }

                        if (options?.Margin != null) {
                            page.Margin(options.Margin.Value, options.MarginUnit);
                        } else {
                            float leftMargin = section.Margins.Left.Value / 20f;
                            float rightMargin = section.Margins.Right.Value / 20f;
                            float topMargin = section.Margins.Top.Value / 20f;
                            float bottomMargin = section.Margins.Bottom.Value / 20f;
                            page.MarginLeft(leftMargin, Unit.Point);
                            page.MarginRight(rightMargin, Unit.Point);
                            page.MarginTop(topMargin, Unit.Point);
                            page.MarginBottom(bottomMargin, Unit.Point);
                        }

                        PageSize size;
                        if (options?.PageSize != null) {
                            size = options.PageSize;
                        } else if (section.PageSettings.PageSize.HasValue) {
                            size = MapToPageSize(section.PageSettings.PageSize.Value);
                        } else if (options?.DefaultPageSize.HasValue == true) {
                            size = MapToPageSize(options.DefaultPageSize.Value);
                        } else {
                            size = PageSizes.A4;
                        }

                        PdfPageOrientation orientation;
                        if (options?.Orientation != null) {
                            orientation = options.Orientation.Value;
                        } else if (section.PageSettings.PageSize.HasValue) {
                            orientation = section.PageSettings.Orientation == W.PageOrientationValues.Landscape ? PdfPageOrientation.Landscape : PdfPageOrientation.Portrait;
                        } else if (options?.DefaultOrientation != null) {
                            orientation = options.DefaultOrientation == W.PageOrientationValues.Landscape ? PdfPageOrientation.Landscape : PdfPageOrientation.Portrait;
                        } else {
                            orientation = PdfPageOrientation.Portrait;
                        }

                        if (orientation == PdfPageOrientation.Landscape) {
                            size = size.Landscape();
                        } else {
                            size = size.Portrait();
                        }

                        page.Size(size);

                        RenderHeader(page, section);
                        RenderFooter(page, section);

                        page.Content().Column(column => {
                            foreach (WordElement element in section.Elements) {
                                if (element is WordParagraph paragraph) {
                                    column.Item().Element(e => RenderParagraph(e, paragraph, GetMarker(paragraph), options));
                                } else if (element is WordTable table) {
                                    column.Item().Element(e => RenderTable(e, table, GetMarker, options));
                                } else if (element is WordImage image) {
                                    column.Item().Element(e => RenderImage(e, image));
                                } else if (element is WordHyperLink link) {
                                    column.Item().Element(e => RenderHyperLink(e, link));
                                } else if (element is WordShape shape) {
                                    column.Item().Element(e => RenderShape(e, shape));
                                }
                            }
                        });
                    });
                }
            });

            return pdf;

            (int Level, string Marker)? GetMarker(WordParagraph paragraph) {
                if (listMarkers.TryGetValue(paragraph, out var value)) {
                    return value;
                }

                return null;
            }

            void RenderElements(ColumnDescriptor column, IEnumerable<WordParagraph> paragraphs, IEnumerable<WordTable> tables) {
                foreach (WordParagraph paragraph in paragraphs) {
                    column.Item().Element(e => RenderParagraph(e, paragraph, GetMarker(paragraph), options));
                }

                foreach (WordTable table in tables) {
                    column.Item().Element(e => RenderTable(e, table, GetMarker, options));
                }
            }

            void RenderHeader(PageDescriptor page, WordSection section) {
                if (section.Header == null) return;
                bool hasContent =
                    (section.Header.Default != null && (section.Header.Default.Paragraphs.Count > 0 || section.Header.Default.Tables.Count > 0)) ||
                    (section.Header.First != null && (section.Header.First.Paragraphs.Count > 0 || section.Header.First.Tables.Count > 0)) ||
                    (section.Header.Even != null && (section.Header.Even.Paragraphs.Count > 0 || section.Header.Even.Tables.Count > 0));
                if (!hasContent) return;

                page.Header().Layers(layers => {
                    if (section.Header.Default != null && (section.Header.Default.Paragraphs.Count > 0 || section.Header.Default.Tables.Count > 0)) {
                        layers.PrimaryLayer().ShowIf(x => (section.Header.First == null || x.PageNumber > 1) && (section.Header.Even == null || x.PageNumber % 2 == 1)).Column(col => {
                            RenderElements(col, section.Header.Default.Paragraphs, section.Header.Default.Tables);
                        });
                    }

                    if (section.Header.First != null && (section.Header.First.Paragraphs.Count > 0 || section.Header.First.Tables.Count > 0)) {
                        layers.Layer().ShowIf(x => x.PageNumber == 1).Column(col => {
                            RenderElements(col, section.Header.First.Paragraphs, section.Header.First.Tables);
                        });
                    }

                    if (section.Header.Even != null && (section.Header.Even.Paragraphs.Count > 0 || section.Header.Even.Tables.Count > 0)) {
                        layers.Layer().ShowIf(x => x.PageNumber % 2 == 0 && x.PageNumber > 1).Column(col => {
                            RenderElements(col, section.Header.Even.Paragraphs, section.Header.Even.Tables);
                        });
                    }
                });
            }

            void RenderFooter(PageDescriptor page, WordSection section) {
                if (section.Footer == null) return;
                bool hasContent =
                    (section.Footer.Default != null && (section.Footer.Default.Paragraphs.Count > 0 || section.Footer.Default.Tables.Count > 0)) ||
                    (section.Footer.First != null && (section.Footer.First.Paragraphs.Count > 0 || section.Footer.First.Tables.Count > 0)) ||
                    (section.Footer.Even != null && (section.Footer.Even.Paragraphs.Count > 0 || section.Footer.Even.Tables.Count > 0));
                if (!hasContent) return;

                page.Footer().Layers(layers => {
                    if (section.Footer.Default != null && (section.Footer.Default.Paragraphs.Count > 0 || section.Footer.Default.Tables.Count > 0)) {
                        layers.PrimaryLayer().ShowIf(x => (section.Footer.First == null || x.PageNumber > 1) && (section.Footer.Even == null || x.PageNumber % 2 == 1)).Column(col => {
                            RenderElements(col, section.Footer.Default.Paragraphs, section.Footer.Default.Tables);
                        });
                    }

                    if (section.Footer.First != null && (section.Footer.First.Paragraphs.Count > 0 || section.Footer.First.Tables.Count > 0)) {
                        layers.Layer().ShowIf(x => x.PageNumber == 1).Column(col => {
                            RenderElements(col, section.Footer.First.Paragraphs, section.Footer.First.Tables);
                        });
                    }

                    if (section.Footer.Even != null && (section.Footer.Even.Paragraphs.Count > 0 || section.Footer.Even.Tables.Count > 0)) {
                        layers.Layer().ShowIf(x => x.PageNumber % 2 == 0 && x.PageNumber > 1).Column(col => {
                            RenderElements(col, section.Footer.Even.Paragraphs, section.Footer.Even.Tables);
                        });
                    }
                });
            }

        }
    }
}