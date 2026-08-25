using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfProductionWorkflowTestSupport {
    internal static byte[] CreatePdf(params string[] pageTexts) {
        if (pageTexts.Length == 0) throw new ArgumentException("At least one page is required.", nameof(pageTexts));
        PdfDocument document = PdfDocument.Create();
        for (int index = 0; index < pageTexts.Length; index++) {
            if (index > 0) document.PageBreak();
            string text = pageTexts[index];
            document.Paragraph(paragraph => paragraph.Text(text));
        }
        return document.ToBytes();
    }

    internal static string[] ReadPageTexts(byte[] pdf) => PdfReadDocument.Open(pdf).Pages
        .Select(static page => Normalize(page.ExtractText()))
        .ToArray();

    internal static string Normalize(string value) => string.Concat(value.Where(static character => !char.IsWhiteSpace(character)));
}
