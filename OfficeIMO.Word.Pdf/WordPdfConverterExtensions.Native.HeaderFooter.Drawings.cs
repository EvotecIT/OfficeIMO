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
        private static IReadOnlyList<NativeHeaderFooterImage> GetNativeHeaderFooterImages(WordHeaderFooter? headerFooter, WordToPdfOptions? options, string source) {
            if (headerFooter == null) {
                return Array.Empty<NativeHeaderFooterImage>();
            }

            var images = new List<NativeHeaderFooterImage>();
            foreach (WordElement element in CollapseNativeParagraphElements(headerFooter.Elements)) {
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

        private static void AddNativeHeaderFooterTableImages(List<NativeHeaderFooterImage> images, WordTable table, WordToPdfOptions? options, string source) {
            foreach (WordTableRow row in table.Rows) {
                IReadOnlyList<WordTableCell> cells = row.Cells;
                if (cells.Count == 1) {
                    foreach (WordParagraph paragraph in EnumerateNativeTableCellParagraphs(cells[0])) {
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

                    foreach (WordParagraph paragraph in EnumerateNativeTableCellParagraphs(cells[cellIndex])) {
                        AddNativeHeaderFooterParagraphImage(images, paragraph, align, options, source);
                    }
                }
            }
        }

        private static void AddNativeHeaderFooterParagraphImage(List<NativeHeaderFooterImage> images, WordParagraph paragraph, PdfCore.PdfAlign? alignOverride, WordToPdfOptions? options, string source) {
            PdfCore.PdfAlign align = alignOverride ?? ResolveNativeParagraphAlign(paragraph, allowJustify: false);
            int imageLimit = options?.MaxImagesPerParagraph ?? 1_000;
            if (imageLimit <= 0) throw new ArgumentOutOfRangeException(nameof(WordToPdfOptions.MaxImagesPerParagraph));
            int imageCount = 0;
            void AddImage(WordImage image) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (++imageCount > imageLimit)
                    throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
                AddNativeHeaderFooterImage(images, image, align, options, source);
            }
            foreach (WordImage image in EnumerateNativeParagraphImages(paragraph, options?.CancellationToken ?? default))
                AddImage(image);
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
                    foreach (WordParagraph paragraph in EnumerateNativeTableCellParagraphs(cells[0])) {
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

                    foreach (WordParagraph paragraph in EnumerateNativeTableCellParagraphs(cells[cellIndex])) {
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

    }
}
