using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Fonts;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class MigraDocPdfGenerator {
    private static int _configured;

    internal static byte[] Generate(PdfBenchmarkScenario scenario) {
        Configure();
        var document = new Document { Info = { Title = scenario.Name } };
        document.Styles[StyleNames.Normal]!.Font.Name = "Carlito";
        document.Styles[StyleNames.Normal]!.Font.Size = 9;

        for (int page = 1; page <= scenario.PageCount; page++) {
            Section section = document.AddSection();
            section.PageSetup.PageFormat = PageFormat.A4;
            section.PageSetup.TopMargin = Unit.FromPoint(32);
            section.PageSetup.BottomMargin = Unit.FromPoint(32);
            section.PageSetup.LeftMargin = Unit.FromPoint(32);
            section.PageSetup.RightMargin = Unit.FromPoint(32);

            Paragraph heading = section.AddParagraph(scenario.PageTitle(page));
            heading.Format.Font.Size = 18;
            heading.Format.Font.Bold = true;
            heading.Format.SpaceAfter = Unit.FromPoint(8);
            for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++) {
                Paragraph narrative = section.AddParagraph(scenario.Narrative(page, paragraph));
                narrative.Format.SpaceAfter = Unit.FromPoint(5);
            }

            Table table = section.AddTable();
            table.Borders.Width = 0.5;
            table.AddColumn(Unit.FromCentimeter(6));
            table.AddColumn(Unit.FromCentimeter(5));
            table.AddColumn(Unit.FromCentimeter(3));
            table.AddColumn(Unit.FromCentimeter(3));
            foreach (string[] sourceRow in scenario.TableRows(page)) {
                Row row = table.AddRow();
                for (int column = 0; column < sourceRow.Length; column++) {
                    row.Cells[column].AddParagraph(sourceRow[column]);
                }
            }
        }

        var renderer = new MigraDoc.Rendering.PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        using var output = new MemoryStream();
        renderer.PdfDocument.Save(output, closeStream: false);
        return output.ToArray();
    }

    private static void Configure() {
        if (Interlocked.Exchange(ref _configured, 1) == 0) {
            GlobalFontSettings.FontResolver = new BenchmarkFontResolver();
        }
    }

    private sealed class BenchmarkFontResolver : IFontResolver {
        public byte[]? GetFont(string faceName) =>
            string.Equals(faceName, "Carlito", StringComparison.OrdinalIgnoreCase)
                ? PdfBenchmarkAssets.CarlitoRegular
                : null;

        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic) =>
            new("Carlito");
    }
}
