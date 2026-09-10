using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.OneNote;

namespace OfficeIMO.ConversionConsistency;

internal static partial class FixtureCorpus {
    internal static async Task CreateAsync(string repository, string output) {
        if (Directory.Exists(output) && Directory.EnumerateFileSystemEntries(output).Any())
            throw new IOException("Fixture output must be empty.");
        Directory.CreateDirectory(output);
        const string family = "Consistency Sans";
        using (var word = WordDocument.Create()) {
            word.Sections[0].PageSettings.PageSize = WordPageSize.Letter;
            for (int page = 1; page <= 2; page++) {
                if (page > 1) word.AddPageBreak();
                var heading = word.AddParagraph("SERVICE REVIEW " + page);
                heading.FontFamily = family; heading.FontSize = 18;
                var body = word.AddParagraph("Availability 99.95 percent. Reviewed actions: 12. Completed actions: 9.");
                body.FontFamily = family; body.FontSize = 12;
            }
            word.Save(Path.Combine(output, "service-review.docx"));
            word.ToOpenDocumentResult().Value.Save(Path.Combine(output, "service-review.odt"));
        }
        using (var excel = ExcelDocument.Create()) {
            for (int page = 1; page <= 2; page++) {
                var sheet = excel.AddWorksheet("Review " + page);
                sheet.CellValue(1, 1, "SERVICE REVIEW " + page);
                sheet.CellValue(2, 1, "Availability"); sheet.CellValue(2, 2, 99.95D);
                sheet.CellValue(3, 1, "Completed"); sheet.CellValue(3, 2, 9);
                sheet.SetColumnWidth(1, 28); sheet.SetColumnWidth(2, 14);
                for (int row = 1; row <= 3; row++) {
                    sheet.SetRowHeight(row, 24);
                    for (int column = 1; column <= 2; column++) sheet.CellFontName(row, column, family);
                }
            }
            excel.Save(Path.Combine(output, "service-review.xlsx"));
            excel.ToOpenDocumentResult().Value.Save(Path.Combine(output, "service-review.ods"));
        }
        using (var presentation = PowerPointPresentation.Create()) {
            presentation.SlideSize.WidthCm = 12.7;
            presentation.SlideSize.HeightCm = 8.4666666667;
            for (int page = 1; page <= 2; page++) {
                var slide = presentation.AddSlide();
                var heading = slide.AddTextBoxPoints("SERVICE REVIEW " + page, 18, 18, 320, 30);
                heading.FontName = family; heading.FontSize = 18;
                var body = slide.AddTextBoxPoints("Completed actions: 9 of 12", 18, 66, 320, 30);
                body.FontName = family; body.FontSize = 12;
            }
            presentation.Save(Path.Combine(output, "service-review.pptx"));
            presentation.ToOpenDocumentResult().Value.Save(Path.Combine(output, "service-review.odp"));
        }
        var section = new OneNoteSection { Name = "Service review" };
        for (int number = 1; number <= 2; number++) {
            var page = new OneNotePage { Title = "SERVICE REVIEW " + number };
            var paragraph = new OneNoteParagraph();
            paragraph.Runs.Add(new OneNoteTextRun { Text = "Completed actions: 9 of 12", Style = { FontFamily = family, FontSize = 12 } });
            page.DirectContent.Add(paragraph);
            section.Pages.Add(page);
        }
        section.Save(Path.Combine(output, "service-review.one"));
        var cases = new List<ConsistencyCase>();
        foreach (var item in new[] { ("docx",816,1056), ("odt",816,1056), ("xlsx",304,96), ("ods",304,96), ("pptx",480,320), ("odp",480,320), ("one",816,1056) }) {
            cases.Add(new ConsistencyCase {
                Id = "native-" + item.Item1, Format = item.Item1,
                Source = Path.GetRelativePath(repository, Path.Combine(output, "service-review." + item.Item1)).Replace('\\','/'),
                Evidence = "Two authored pages or sheets with distinct SERVICE REVIEW labels; source generation is in FixtureCorpus.cs. This suite detects route differences and is not an approval baseline.",
                ComparePdfPixels = item.Item1 is not ("xlsx" or "ods"),
                // Tight worksheet crops contain a high proportion of 11px text. Font rasterizers
                // differ at glyph edges; separate baseline assertions prevent accepting a layout shift.
                VisualTolerance = item.Item1 is "xlsx" or "ods"
                    ? new PixelTolerance { DifferentRatio = 0.04D, MeanAbsoluteError = 4.5D } : new(),
                RequireSearchablePdf = item.Item1 != "one",
                Limitations = item.Item1 is "xlsx" or "ods"
                    ? new() { "Worksheet images use tight content bounds; PDF uses a paper page. PDF page dimensions and text are checked, but PDF pixels are not compared with worksheet pixels." }
                    : item.Item1 == "one" ? new() {
                        "The native OneNote visual PDF contains raster pages and has no searchable text layer. SVG labels and PDF pixels are checked.",
                        "The OneNote title requests bold. This regular-face fixture profile deliberately exercises synthetic bold and accepts the reported substitution to its pinned regular face."
                    } : new(),
                AllowedDiagnostics = item.Item1 == "ods" ? new() { "ODF_IMAGE_COLUMN_LAYOUT_APPROXIMATED", "ODF_IMAGE_CELL_STYLES_APPROXIMATED" }
                    : item.Item1 == "one" ? new() { "IMAGE_FONT_SUBSTITUTED" }
                    : item.Item1 == "odp" ? new() { "ODF_IMAGE_MASTERS_LAYOUTS_APPROXIMATED" } : new(),
                Pages = Enumerable.Range(1,2).Select(number => new PageExpectation {
                    Width = item.Item2, Height = item.Item3,
                    PdfWidth = item.Item1 is "xlsx" or "ods" ? 816 : null,
                    PdfHeight = item.Item1 is "xlsx" or "ods" ? 1056 : null,
                    SvgTextPositions = item.Item1 is "xlsx" or "ods" ? new() {
                        new("SERVICE REVIEW " + number, 3, 27.74, 0.1),
                        new("Availability", 3, 59.74, 0.1),
                        new("Completed", 3, 91.74, 0.1)
                    } : new(),
                    Text = new() { "SERVICE REVIEW " + number }
                }).ToList()
            });
        }
        AddGroupedPowerPoint(repository, output, family, cases);
        await AddOtherFormatsAsync(repository, output, family, cases);
        GateJson.Write(Path.Combine(output, "suite.json"), new ConsistencySuite {
            FontPath = "OfficeIMO.Drawing.Tests/TestAssets/OpenSans-Regular.woff2", FontFamily = family, Cases = cases
        });
        Console.WriteLine("Prepared native fixtures: " + output);
    }
}
