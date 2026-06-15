using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    internal static PdfCore.PdfDocument Convert(RtfDocument document, RtfPdfSaveOptions? options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        RtfPdfSaveOptions normalized = (options ?? new RtfPdfSaveOptions()).Normalize();
        PdfCore.PdfOptions pdfOptions = normalized.PdfOptions ?? new PdfCore.PdfOptions();
        ApplyPageSetup(document, document.PageSetup, pdfOptions);
        if (document.Sections.Count > 0) {
            ApplyPageSetup(document, document.Sections[0].PageSetup, pdfOptions);
        }

        ApplyHeaderFooters(document, pdfOptions, normalized);

        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(pdfOptions);
        ApplyMetadata(document, pdf, normalized);
        PdfRenderState state = new PdfRenderState(document);

        RenderDocumentBlocks(document, pdf, normalized, state, pdfOptions);

        RenderNotes(document, pdf, normalized, state);
        return pdf;
    }

    private static void RenderDocumentBlocks(RtfDocument document, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state, PdfCore.PdfOptions pdfOptions) {
        if (document.Sections.Count == 0) {
            RenderBlocks(document, document.Blocks, pdf, options, state);
            return;
        }

        for (int index = 0; index < document.Sections.Count; index++) {
            RtfSection section = document.Sections[index];
            if (index == 0) {
                RenderBlocks(document, section.Blocks, pdf, options, state);
                continue;
            }

            if (!StartsNewPdfPage(section.BreakKind)) {
                RenderBlocks(document, section.Blocks, pdf, options, state);
                continue;
            }

            pdf.Section(page => {
                ApplyPageSetup(document, section.PageSetup, page, pdfOptions);
                RenderBlocks(document, section.Blocks, pdf, options, state);

                while (index + 1 < document.Sections.Count && !StartsNewPdfPage(document.Sections[index + 1].BreakKind)) {
                    index++;
                    RenderBlocks(document, document.Sections[index].Blocks, pdf, options, state);
                }
            });
        }
    }

    private static void RenderBlocks(RtfDocument document, IEnumerable<IRtfBlock> blocks, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        foreach (IRtfBlock block in blocks) {
            RenderBlock(document, block, pdf, options, state);
        }
    }

    private static bool StartsNewPdfPage(RtfSectionBreakKind breakKind) {
        switch (breakKind) {
            case RtfSectionBreakKind.Continuous:
                return false;
            default:
                return true;
        }
    }

    private static void RenderBlock(RtfDocument document, IRtfBlock block, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        switch (block) {
            case RtfParagraph paragraph:
                RenderParagraph(document, paragraph, pdf, options, state);
                break;
            case RtfTable table when options.IncludeTables:
                RenderTable(document, table, pdf, options, state);
                break;
            case RtfImage image:
                RenderImage(image, pdf, options);
                break;
            case RtfObject rtfObject:
                RenderPlainTextBlock(rtfObject.ToPlainText(), pdf);
                break;
            case RtfShape shape:
                RenderPlainTextBlock(shape.ToPlainText(), pdf);
                break;
        }
    }

    private static void RenderPlainTextBlock(string text, PdfCore.PdfDocument pdf) {
        if (!string.IsNullOrEmpty(text)) {
            pdf.Paragraph(paragraph => paragraph.Text(text));
        }
    }
}
