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
        private static List<PdfFootnote> CollectNativeFootnotes(IReadOnlyList<WordElement> elements, Dictionary<long, int> footnoteNumbersById) {
            var footnotes = new List<PdfFootnote>();
            foreach (WordElement element in elements) {
                CollectNativeFootnotes(element, footnotes, footnoteNumbersById, structuredDocumentTagDepth: 0);
            }

            return footnotes;
        }

        private static void CollectNativeFootnotes(
            WordElement element,
            List<PdfFootnote> footnotes,
            Dictionary<long, int> footnoteNumbersById,
            int structuredDocumentTagDepth) {
            switch (element) {
                case WordFootNote footNote:
                    AddNativeFootnote(footNote, footnotes, footnoteNumbersById);
                    break;
                case WordEndNote endNote:
                    AddNativeEndnote(endNote, footnotes, footnoteNumbersById);
                    break;
                case WordParagraph paragraph:
                    WordFootNote? paragraphFootnote = paragraph.FootNote;
                    if (paragraphFootnote != null) {
                        AddNativeFootnote(paragraphFootnote, footnotes, footnoteNumbersById);
                    }

                    WordEndNote? paragraphEndnote = paragraph.EndNote;
                    if (paragraphEndnote != null) {
                        AddNativeEndnote(paragraphEndnote, footnotes, footnoteNumbersById);
                    }

                    foreach (WordParagraph run in paragraph.GetRuns()) {
                        WordFootNote? runFootnote = run.FootNote;
                        if (runFootnote != null) {
                            AddNativeFootnote(runFootnote, footnotes, footnoteNumbersById);
                        }

                        WordEndNote? runEndnote = run.EndNote;
                        if (runEndnote != null) {
                            AddNativeEndnote(runEndnote, footnotes, footnoteNumbersById);
                        }
                    }

                    break;
                case WordTable table:
                    foreach (WordTable currentTable in EnumerateNativeTableTree(table)) {
                        foreach (WordTableRow row in currentTable.Rows) {
                            foreach (WordTableCell cell in row.Cells) {
                                foreach (WordParagraph paragraph in cell.Paragraphs) {
                                    CollectNativeFootnotes(paragraph, footnotes, footnoteNumbersById, structuredDocumentTagDepth);
                                }
                            }
                        }
                    }

                    break;
                case WordCoverPage coverPage:
                    EnsureNativeStructuredDocumentTagDepth(structuredDocumentTagDepth);
                    foreach (WordElement coverElement in GetNativeStructuredBlockElements(coverPage.Document, coverPage.SdtBlock)) {
                        CollectNativeFootnotes(coverElement, footnotes, footnoteNumbersById, structuredDocumentTagDepth + 1);
                    }

                    break;
                case WordStructuredDocumentTag structuredDocumentTag:
                    EnsureNativeStructuredDocumentTagDepth(structuredDocumentTagDepth);
                    foreach (WordElement structuredElement in GetNativeStructuredBlockElements(structuredDocumentTag.Document, structuredDocumentTag.SdtBlock)) {
                        CollectNativeFootnotes(structuredElement, footnotes, footnoteNumbersById, structuredDocumentTagDepth + 1);
                    }

                    break;
            }
        }

        private static void EnsureNativeStructuredDocumentTagDepth(int depth) {
            if (depth >= MaximumNativeStructuredDocumentTagDepth) {
                throw new InvalidDataException(
                    $"Structured document tag nesting exceeds the supported limit of {MaximumNativeStructuredDocumentTagDepth} levels.");
            }
        }

        private static void AddNativeFootnote(WordFootNote footNote, List<PdfFootnote> footnotes, Dictionary<long, int> footnoteNumbersById) {
            long? referenceId = footNote.ReferenceId;
            if (!referenceId.HasValue || referenceId.Value == 0) {
                return;
            }

            long key = GetNativeFootnoteKey(referenceId.Value);
            if (footnoteNumbersById.ContainsKey(key)) {
                return;
            }

            int number = footnoteNumbersById.Keys.Count(key => key > 0) + 1;
            footnoteNumbersById[key] = number;
            footnotes.Add(new PdfFootnote {
                Number = number,
                Text = GetNativeFootnoteText(footNote)
            });
        }

        private static void AddNativeEndnote(WordEndNote endNote, List<PdfFootnote> footnotes, Dictionary<long, int> footnoteNumbersById) {
            long? referenceId = endNote.ReferenceId;
            if (!referenceId.HasValue || referenceId.Value == 0) {
                return;
            }

            long key = GetNativeEndnoteKey(referenceId.Value);
            if (footnoteNumbersById.ContainsKey(key)) {
                return;
            }

            int number = footnoteNumbersById.Keys.Count(key => key < 0) + 1;
            footnoteNumbersById[key] = number;
            footnotes.Add(new PdfFootnote {
                Number = number,
                Text = GetNativeEndnoteText(endNote)
            });
        }

        private static string GetNativeFootnoteText(WordFootNote footNote) {
            var parts = new List<string>();
            foreach (WordParagraph paragraph in footNote.Paragraphs ?? Enumerable.Empty<WordParagraph>()) {
                if (!string.IsNullOrWhiteSpace(paragraph.Text)) {
                    parts.Add(paragraph.Text);
                }
            }

            return string.Join(" ", parts);
        }

        private static string GetNativeEndnoteText(WordEndNote endNote) {
            var parts = new List<string>();
            foreach (WordParagraph paragraph in endNote.Paragraphs ?? Enumerable.Empty<WordParagraph>()) {
                if (!string.IsNullOrWhiteSpace(paragraph.Text)) {
                    parts.Add(paragraph.Text);
                }
            }

            return string.Join(" ", parts);
        }

        private static IReadOnlyList<int> GetNativeFootnoteNumbersForElement(IReadOnlyList<WordElement> elements, int index, Dictionary<long, int> footnoteNumbersById) {
            var numbers = new List<int>();
            for (int i = index + 1; i < elements.Count && (elements[i] is WordFootNote || elements[i] is WordEndNote); i++) {
                long? key = GetNativeNoteKey(elements[i]);
                if (key.HasValue && footnoteNumbersById.TryGetValue(key.Value, out int number)) {
                    numbers.Add(number);
                }
            }

            return numbers;
        }

        private static List<int> GetNativeParagraphFootnoteNumbers(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, IReadOnlyList<int> followingFootnoteNumbers, Dictionary<long, int> footnoteNumbersById) {
            var numbers = new List<int>(followingFootnoteNumbers);
            AddNativeParagraphFootnoteNumber(paragraph, numbers, footnoteNumbersById);
            foreach (WordParagraph run in runs) {
                AddNativeParagraphFootnoteNumber(run, numbers, footnoteNumbersById);
            }

            return numbers.Distinct().ToList();
        }

        private static void AddNativeParagraphFootnoteNumber(WordParagraph paragraph, List<int> numbers, Dictionary<long, int> footnoteNumbersById) {
            WordFootNote? footNote = paragraph.FootNote;
            long? footnoteKey = footNote?.ReferenceId.HasValue == true && footNote.ReferenceId.Value != 0 ? GetNativeFootnoteKey(footNote.ReferenceId.Value) : null;
            if (footnoteKey.HasValue && footnoteNumbersById.TryGetValue(footnoteKey.Value, out int number)) {
                numbers.Add(number);
            }

            WordEndNote? endNote = paragraph.EndNote;
            long? endnoteKey = endNote?.ReferenceId.HasValue == true && endNote.ReferenceId.Value != 0 ? GetNativeEndnoteKey(endNote.ReferenceId.Value) : null;
            if (endnoteKey.HasValue && footnoteNumbersById.TryGetValue(endnoteKey.Value, out number)) {
                numbers.Add(number);
            }
        }

        private static long? GetNativeNoteKey(WordElement element) {
            switch (element) {
                case WordFootNote footNote when footNote.ReferenceId.HasValue && footNote.ReferenceId.Value != 0:
                    return GetNativeFootnoteKey(footNote.ReferenceId.Value);
                case WordEndNote endNote when endNote.ReferenceId.HasValue && endNote.ReferenceId.Value != 0:
                    return GetNativeEndnoteKey(endNote.ReferenceId.Value);
                default:
                    return null;
            }
        }

        private static long GetNativeFootnoteKey(long referenceId) => referenceId;

        private static long GetNativeEndnoteKey(long referenceId) => -referenceId;

        private static void RenderNativeFootnotes(INativePdfFlow pdf, IReadOnlyList<PdfFootnote> footnotes) {
            if (footnotes.Count == 0) {
                return;
            }

            pdf.HR(thickness: 0.5, color: PdfCore.PdfColor.LightGray, spacingBefore: 8, spacingAfter: 4);
            foreach (PdfFootnote footnote in footnotes) {
                pdf.Paragraph(builder => {
                    builder.Baseline(PdfCore.PdfTextBaseline.Superscript);
                    builder.Text(footnote.Number.ToString(CultureInfo.InvariantCulture));
                    builder.Baseline(PdfCore.PdfTextBaseline.Normal);
                    if (!string.IsNullOrWhiteSpace(footnote.Text)) {
                        builder.Text(" ");
                        builder.Text(NormalizeNativeDirectText(footnote.Text));
                    }
                });
            }
        }

        private static void RenderNativeImage(INativePdfFlow pdf, WordImage image, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, PdfSaveOptions? options = null, string source = "body image") {
            if (image == null) {
                return;
            }

            if (!TryGetNativeBodyImageBytes(image, options, source, out byte[] bytes)) {
                return;
            }

            if (!TryPrepareNativePdfImageBytes(bytes, out byte[] preparedBytes, out string? unsupportedReason)) {
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "NativeBodyImageUnsupported",
                        source,
                        "Word image was not exported because the shared PDF raster pipeline could not prepare it. " + unsupportedReason);
                }

                return;
            }

            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            pdf.Image(preparedBytes, width, height, align);
        }

        private static bool TryGetNativeBodyImageBytes(WordImage image, PdfSaveOptions? options, string source, out byte[] bytes) {
            try {
                bytes = ImageEmbedder.GetImageBytes(image);
                return true;
            } catch (InvalidOperationException ex) {
                bytes = System.Array.Empty<byte>();
                if (options != null) {
                    AddNativeExportWarning(
                        options,
                        "NativeBodyImageUnavailable",
                        source,
                        "Word image was not exported because the image bytes could not be extracted. " + ex.Message);
                }

                return false;
            }
        }

        private static bool TryPrepareNativePdfImageBytes(
            byte[] bytes,
            out byte[] preparedBytes,
            out string? unsupportedReason) =>
            PdfCore.PdfDocument.TryPrepareImageBytes(
                bytes,
                out preparedBytes,
                out _,
                out _,
                out unsupportedReason);

    }
}
