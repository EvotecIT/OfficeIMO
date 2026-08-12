using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class QuestPdfGenerator {
    private static int _configured;

    internal static byte[] Generate(PdfBenchmarkScenario scenario) {
        Configure();
        return Document.Create(container => {
            for (int pageNumber = 1; pageNumber <= scenario.PageCount; pageNumber++) {
                int page = pageNumber;
                container.Page(descriptor => {
                    descriptor.Size(QuestPDF.Helpers.PageSizes.A4);
                    descriptor.Margin(32);
                    descriptor.DefaultTextStyle(style => style.FontFamily("Carlito").FontSize(9));
                    descriptor.Content().Column(column => {
                        column.Spacing(8);
                        column.Item().Text(scenario.PageTitle(page)).Bold().FontSize(18);
                        for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++) {
                            column.Item().Text(scenario.Narrative(page, paragraph));
                        }
                        column.Item().Table(table => {
                            table.ColumnsDefinition(columns => {
                                columns.RelativeColumn(2);
                                columns.RelativeColumn(2);
                                columns.RelativeColumn(1);
                                columns.RelativeColumn(1);
                            });
                            foreach (string[] row in scenario.TableRows(page)) {
                                foreach (string cell in row) {
                                    table.Cell().BorderBottom(0.5f).Padding(3).Text(cell);
                                }
                            }
                        });
                    });
                });
            }
        }).GeneratePdf();
    }

    private static void Configure() {
        if (Interlocked.Exchange(ref _configured, 1) != 0) {
            return;
        }

        QuestPDF.Settings.License = LicenseType.Community;
        using var font = new MemoryStream(PdfBenchmarkAssets.CarlitoRegular, writable: false);
        QuestPDF.Drawing.FontManager.RegisterFont(font);
    }
}
