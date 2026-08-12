using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class ITextPdfGenerator {
    internal static byte[] Generate(PdfBenchmarkScenario scenario) {
        using var output = new MemoryStream();
        var writer = new PdfWriter(output, new WriterProperties().SetCompressionLevel(6));
        var pdf = new iText.Kernel.Pdf.PdfDocument(writer);
        var document = new Document(pdf, iText.Kernel.Geom.PageSize.A4);
        document.SetMargins(32, 32, 32, 32);
        PdfFont font = PdfFontFactory.CreateFont(
            PdfBenchmarkAssets.CarlitoRegular,
            PdfEncodings.IDENTITY_H,
            PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
        document.SetFont(font).SetFontSize(9);

        for (int page = 1; page <= scenario.PageCount; page++) {
            document.Add(new Paragraph(scenario.PageTitle(page)).SetFontSize(18));
            for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++) {
                document.Add(new Paragraph(scenario.Narrative(page, paragraph)));
            }

            var table = new Table(UnitValue.CreatePercentArray(new float[] { 2, 2, 1, 1 }))
                .UseAllAvailableWidth();
            foreach (string[] sourceRow in scenario.TableRows(page)) {
                foreach (string cell in sourceRow) {
                    table.AddCell(new Cell().Add(new Paragraph(cell).SetMargin(0)).SetPadding(3));
                }
            }
            document.Add(table);
            if (page < scenario.PageCount) {
                document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
            }
        }

        document.Close();
        return output.ToArray();
    }
}
