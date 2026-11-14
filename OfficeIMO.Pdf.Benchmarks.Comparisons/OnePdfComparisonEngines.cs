#if NET10_0_OR_GREATER
using System.Text;
using OnePdf.Layout;
using OnePdf.Text;
using OnePdfDocument = OnePdf.PdfDocument;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// OnePDF implementations of the shared read and generation contracts. The packages target
/// net10.0 only, so these lanes exist only in the net10.0 build of the benchmark runner.
/// </summary>
internal static class OnePdfComparisonEngines {
    private static readonly PdfTextOptions SequentialText = new() { MaxDegreeOfParallelism = 1 };
    private static readonly PdfTextOptions DefaultText = new();

    /// <summary>Whole-document extraction on one thread, like the other readers in the matrix.</summary>
    internal static PdfReadObservation Read(byte[] pdf) => Read(pdf, SequentialText);

    /// <summary>Whole-document extraction with the library's default page parallelism (min(4, cores)).</summary>
    internal static PdfReadObservation ReadParallel(byte[] pdf) => Read(pdf, DefaultText);

    internal static IReadOnlyList<string> ExtractPages(byte[] pdf) {
        using OnePdfDocument document = OnePdfDocument.Open(pdf);
        return document.ExtractText(SequentialText).Select(static page => page.Text).ToArray();
    }

    private static PdfReadObservation Read(byte[] pdf, PdfTextOptions options) {
        using OnePdfDocument document = OnePdfDocument.Open(pdf);
        IReadOnlyList<PdfPageText> pages = document.ExtractText(options);
        return PdfBenchmarkValidation.Observe(document.PageCount, JoinPages(pages));
    }

    private static string JoinPages(IReadOnlyList<PdfPageText> pages) {
        var text = new StringBuilder();
        foreach (PdfPageText page in pages) {
            text.Append(page.Text);
            text.Append('\n');
        }
        return text.ToString();
    }

    /// <summary>Lays out the shared report contract: one page per scenario page, title, narrative and a four-column table.</summary>
    internal static byte[] Generate(PdfBenchmarkScenario scenario) =>
        LayoutDocument.Compose(document => {
            document.Title(scenario.Name);
            document.Fonts.Add(PdfBenchmarkAssets.CarlitoRegular, "Carlito");
            document.Fonts.Add(PdfBenchmarkAssets.CarlitoBold, "Carlito");
            for (int pageNumber = 1; pageNumber <= scenario.PageCount; pageNumber++) {
                int page = pageNumber;
                document.Page(descriptor => {
                    descriptor.Size(PageSize.A4);
                    descriptor.Margin(32);
                    descriptor.Content().Column(column => {
                        column.Spacing(8);
                        column.Item().Heading(1).Text(scenario.PageTitle(page)).FontFamily("Carlito").FontSize(18).Bold();
                        for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++) {
                            column.Item().Text(scenario.Narrative(page, paragraph)).FontFamily("Carlito").FontSize(9);
                        }
                        column.Item().Table(table => {
                            table.Columns(columns => {
                                columns.Relative(2);
                                columns.Relative(2);
                                columns.Relative(1);
                                columns.Relative(1);
                            });
                            table.CellPadding(3);
                            foreach (string[] row in scenario.TableRows(page)) {
                                table.Row(cells => {
                                    foreach (string cell in row) {
                                        cells.Cell().BorderBottom(0.5).Text(cell).FontFamily("Carlito").FontSize(9);
                                    }
                                });
                            }
                        });
                    });
                });
            }
        }).ToBytes();
}
#endif
