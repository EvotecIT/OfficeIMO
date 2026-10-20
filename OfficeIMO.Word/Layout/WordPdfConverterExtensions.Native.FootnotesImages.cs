using System.Collections.Generic;
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
        private static List<PdfFootnote> CollectNativeFootnotes(IReadOnlyList<WordElement> elements, NativeNoteNumbering footnoteNumbersById,
            List<PdfFootnote> documentEndnotes) {
            var footnotes = new List<PdfFootnote>();
            foreach (WordElement element in elements) {
                CollectNativeFootnotes(element, footnotes, footnoteNumbersById, structuredDocumentTagDepth: 0, tableDepth: 0);
            }

            if (!footnoteNumbersById.EndnotesAtDocumentEnd) return footnotes;
            var sectionNotes = new List<PdfFootnote>(footnotes.Count);
            foreach (PdfFootnote note in footnotes) {
                if (note.IsEndnote) documentEndnotes.Add(note);
                else sectionNotes.Add(note);
            }
            return sectionNotes;
        }

        private static void CollectNativeFootnotes(
            WordElement element,
            List<PdfFootnote> footnotes,
            NativeNoteNumbering footnoteNumbersById,
            int structuredDocumentTagDepth,
            int tableDepth) {
            switch (element) {
                case WordFootNote footNote:
                    AddNativeFootnote(footNote, footnotes, footnoteNumbersById);
                    break;
                case WordEndNote endNote:
                    AddNativeEndnote(endNote, footnotes, footnoteNumbersById);
                    break;
                case WordParagraph paragraph:
                    // Read actual reference children: a single Word run can carry several
                    // notes, or a note between text siblings, and SDTs retain this order.
                    foreach (var reference in EnumerateNativeVisibleNoteReferences(paragraph)) {
                        if (reference is W.FootnoteReference foot && foot.Id?.Value is long footId)
                            AddNativeFootnote(paragraph._document, footId, footnotes, footnoteNumbersById);
                        else if (reference is W.EndnoteReference end && end.Id?.Value is long endId)
                            AddNativeEndnote(paragraph._document, endId, footnotes, footnoteNumbersById);
                    }

                    break;
                case WordTable table:
                    EnsureNativeTableDepth(tableDepth);
                    foreach (WordTableRow row in table.Rows) {
                        foreach (WordTableCell cell in row.Cells) {
                            foreach (WordElement cellElement in cell.Elements) {
                                CollectNativeFootnotes(
                                    cellElement,
                                    footnotes,
                                    footnoteNumbersById,
                                    structuredDocumentTagDepth,
                                    tableDepth + 1);
                            }
                        }
                    }

                    break;
                case WordCoverPage coverPage:
                    EnsureNativeStructuredDocumentTagDepth(structuredDocumentTagDepth);
                    foreach (WordElement coverElement in GetNativeStructuredBlockElements(coverPage.Document, coverPage.SdtBlock)) {
                        CollectNativeFootnotes(coverElement, footnotes, footnoteNumbersById, structuredDocumentTagDepth + 1, tableDepth);
                    }

                    break;
                case WordStructuredDocumentTag structuredDocumentTag:
                    EnsureNativeStructuredDocumentTagDepth(structuredDocumentTagDepth);
                    if (TryGetNativeStructuredBlockPropertyValue(structuredDocumentTag.Document, structuredDocumentTag.SdtBlock, out _)) break;
                    foreach (WordElement structuredElement in GetNativeStructuredBlockElements(structuredDocumentTag.Document, structuredDocumentTag.SdtBlock)) {
                        CollectNativeFootnotes(structuredElement, footnotes, footnoteNumbersById, structuredDocumentTagDepth + 1, tableDepth);
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

        private static void AddNativeFootnote(WordFootNote footNote, List<PdfFootnote> footnotes, NativeNoteNumbering footnoteNumbersById) {
            long? referenceId = footNote.ReferenceId;
            if (!referenceId.HasValue) {
                return;
            }

            AddNativeFootnote(footNote.Paragraphs?.FirstOrDefault()?._document, referenceId.Value, footnotes, footnoteNumbersById);
        }

        private static void AddNativeFootnote(WordDocument? document, long referenceId,
            List<PdfFootnote> footnotes, NativeNoteNumbering footnoteNumbersById) {
            NativeNoteKey key = GetNativeFootnoteKey(referenceId);
            if (footnoteNumbersById.ContainsKey(key)) return;
            W.Footnote? source = WordNoteResolver.FindFootnote(document?._wordprocessingDocument.MainDocumentPart, referenceId);
            if (source == null) return;
            footnotes.Add(new PdfFootnote {
                Label = footnoteNumbersById.Add(key, endnote: false),
                Text = string.Join(" ", source.Descendants<W.Paragraph>().Select(WordParagraph.ReadVisibleText))
            });
        }

        private static void AddNativeEndnote(WordEndNote endNote, List<PdfFootnote> footnotes, NativeNoteNumbering footnoteNumbersById) {
            long? referenceId = endNote.ReferenceId;
            if (!referenceId.HasValue) {
                return;
            }

            AddNativeEndnote(endNote.Paragraphs?.FirstOrDefault()?._document, referenceId.Value, footnotes, footnoteNumbersById);
        }

        private static void AddNativeEndnote(WordDocument? document, long referenceId,
            List<PdfFootnote> footnotes, NativeNoteNumbering footnoteNumbersById) {
            NativeNoteKey key = GetNativeEndnoteKey(referenceId);
            if (footnoteNumbersById.ContainsKey(key)) return;
            W.Endnote? source = WordNoteResolver.FindEndnote(document?._wordprocessingDocument.MainDocumentPart, referenceId);
            if (source == null) return;
            footnotes.Add(new PdfFootnote {
                Label = footnoteNumbersById.Add(key, endnote: true),
                Text = string.Join(" ", source.Descendants<W.Paragraph>().Select(WordParagraph.ReadVisibleText)),
                IsEndnote = true
            });
        }

        private static IReadOnlyList<int> GetNativeFootnoteNumbersForElement(IReadOnlyList<WordElement> elements, int index, NativeNoteNumbering footnoteNumbersById, HashSet<NativeNoteKey>? seenKeys = null) {
            var numbers = new List<int>();
            seenKeys ??= new HashSet<NativeNoteKey>();
            for (int i = index + 1; i < elements.Count && (elements[i] is WordFootNote || elements[i] is WordEndNote); i++) {
                NativeNoteKey? key = GetNativeNoteKey(elements[i]);
                if (key.HasValue && footnoteNumbersById.TryGetValue(key.Value, out int number) && seenKeys.Add(key.Value)) {
                    numbers.Add(number);
                }
            }

            return numbers;
        }

        private static List<int> GetNativeParagraphFootnoteNumbers(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, IReadOnlyList<int> followingFootnoteNumbers, NativeNoteNumbering footnoteNumbersById) {
            var numbers = new List<int>();
            var seenKeys = new HashSet<NativeNoteKey>();
            AddNativeParagraphFootnoteNumber(paragraph, numbers, footnoteNumbersById, seenKeys);
            foreach (WordParagraph run in runs) {
                AddNativeParagraphFootnoteNumber(run, numbers, footnoteNumbersById, seenKeys);
            }

            // Paragraph and traversal wrappers can expose the same reference twice.
            // Tokens retain identity even when displayed labels repeat.
            var unmatched = numbers.GroupBy(number => number).ToDictionary(group => group.Key, group => group.Count());
            foreach (int number in followingFootnoteNumbers) {
                if (unmatched.TryGetValue(number, out int count) && count > 0) unmatched[number] = count - 1;
                else numbers.Add(number);
            }
            return numbers;
        }

        private static void AddNativeParagraphFootnoteNumber(WordParagraph paragraph, List<int> numbers, NativeNoteNumbering footnoteNumbersById, HashSet<NativeNoteKey> seenKeys) {
            if (IsNativeHiddenTextRun(paragraph)) return;
            foreach (var child in paragraph.EnumerateEffectiveRunContent().SelectMany(element => element.Descendants().Prepend(element))) {
                NativeNoteKey? key = child is W.FootnoteReference foot && foot.Id?.Value is long footId
                    ? GetNativeFootnoteKey(footId)
                    : child is W.EndnoteReference end && end.Id?.Value is long endId
                        ? GetNativeEndnoteKey(endId) : null;
                if (key.HasValue && footnoteNumbersById.TryGetValue(key.Value, out int number) && seenKeys.Add(key.Value))
                    numbers.Add(number);
            }
        }

        private static NativeNoteKey? GetNativeNoteKey(WordElement element) {
            switch (element) {
                case WordFootNote footNote when footNote.ReferenceId.HasValue:
                    return GetNativeFootnoteKey(footNote.ReferenceId.Value);
                case WordEndNote endNote when endNote.ReferenceId.HasValue:
                    return GetNativeEndnoteKey(endNote.ReferenceId.Value);
                default:
                    return null;
            }
        }

        private static NativeNoteKey GetNativeFootnoteKey(long referenceId) => new(NativeNoteKind.Footnote, referenceId);

        private static NativeNoteKey GetNativeEndnoteKey(long referenceId) => new(NativeNoteKind.Endnote, referenceId);

        private static void RenderNativeFootnotes(INativePdfFlow pdf, IReadOnlyList<PdfFootnote> footnotes) {
            if (footnotes.Count == 0) {
                return;
            }

            pdf.HR(thickness: 0.5, color: PdfCore.PdfColor.LightGray, spacingBefore: 8, spacingAfter: 4);
            foreach (PdfFootnote footnote in footnotes) {
                pdf.Paragraph(builder => {
                    builder.Baseline(PdfCore.PdfTextBaseline.Superscript);
                    builder.Text(footnote.Label);
                    builder.Baseline(PdfCore.PdfTextBaseline.Normal);
                    if (!string.IsNullOrWhiteSpace(footnote.Text)) {
                        builder.Text(" ");
                        builder.Text(NormalizeNativeDirectText(footnote.Text));
                    }
                });
            }
        }

        private static bool RenderNativeImage(INativePdfFlow pdf, WordImage image, PdfCore.PdfAlign align = PdfCore.PdfAlign.Left, WordToPdfOptions? options = null, string source = "body image", PdfCore.PdfParagraphStyle? anchorStyle = null, PdfCore.PdfPageCanvas? anchoredCanvas = null, NativeObjectParagraphSpacing? paragraphSpacing = null, WordParagraph? anchorParagraph = null, WordSection? anchorSection = null) {
            if (image == null) {
                return false;
            }

            if (!TryPrepareNativeBodyImage(image, options, source, out byte[] preparedBytes)) {
                return false;
            }

            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            if (anchorStyle != null && TryAddNativeAnchoredImage(pdf, image, preparedBytes, width, height,
                    anchorStyle, anchoredCanvas, anchorParagraph, options, source, anchorSection)) return false;
            if (image._Image?.Anchor != null && options != null)
                AddNativeExportWarning(options, "NativeAnchoredImageFlowed", source,
                    "The image's wrap, anchor or transform is outside the fixed-placement contract; it was placed in document flow.");
            RenderNativeFlowObject(pdf, paragraphSpacing, flow => flow.Image(preparedBytes, width, height, align,
                CreateNativeImageStyle(image, options, source)));
            return true;
        }

        private static bool TryPrepareNativeBodyImage(WordImage image, WordToPdfOptions? options, string source, out byte[] bytes,
            string unsupportedCode = "NativeBodyImageUnsupported", System.Threading.CancellationToken cancellationToken = default) {
            try {
                cancellationToken = options?.CancellationToken ?? cancellationToken;
                string? unsupportedReason = null;
                foreach (byte[] candidate in image.GetRenderingImageCandidates(cancellationToken)) {
                    if (TryPrepareNativePdfImageBytes(candidate, out bytes, out unsupportedReason, cancellationToken)) return true;
                }
                bytes = System.Array.Empty<byte>();
                if (options != null) AddNativeExportWarning(options, unsupportedCode, source,
                    "Word image was not exported because the shared PDF image pipeline could not prepare it. " + unsupportedReason);
                return false;
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
            out string? unsupportedReason,
            System.Threading.CancellationToken cancellationToken = default) =>
            PdfCore.PdfDocument.TryPrepareImageBytes(
                bytes,
                cancellationToken,
                out preparedBytes,
                out _,
                out _,
                out unsupportedReason);

    }
}
