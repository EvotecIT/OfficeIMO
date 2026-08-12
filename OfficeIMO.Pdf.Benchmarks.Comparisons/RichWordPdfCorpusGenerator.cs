using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed record RichWordPdfCorpusArtifacts(string DocxPath, string PdfPath, string ConversionReportPath);

internal static class RichWordPdfCorpusGenerator {
    private const int PageCount = 25;

    internal static RichWordPdfCorpusArtifacts Generate(string repositoryRoot, string outputDirectory) {
        Directory.CreateDirectory(outputDirectory);
        string docxPath = Path.Combine(outputDirectory, "officeimo-word-rich-25-page.docx");
        string pdfPath = Path.Combine(outputDirectory, "officeimo-word-rich-25-page.pdf");
        string conversionReportPath = Path.Combine(outputDirectory, "officeimo-word-rich-25-page.conversion.json");
        string imagePath = Path.Combine(repositoryRoot, "Assets", "Images", "Others", "Logo-evotec.png");

        using (WordDocument document = WordDocument.Create(docxPath)) {
            document.AddHeadersAndFooters();
            WordHeader? header = document.Header?.Default;
            WordFooter? footer = document.Footer?.Default;
            if (header == null || footer == null) {
                throw new InvalidDataException("Word fixture did not create default header and footer parts.");
            }
            header.AddParagraph("OfficeIMO rich interoperability corpus");
            footer.AddParagraph("Generated benchmark fixture");

            for (int page = 1; page <= PageCount; page++) {
                document.AddParagraph($"RICH WORD CORPUS PAGE {page:D3}")
                    .SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph(
                    "Deterministic operational evidence with Unicode: Zażółć gęślą jaźń.");

                WordTable table = document.AddTable(6, 4);
                string[] headings = { "Account", "Owner", "Amount", "Status" };
                for (int column = 0; column < headings.Length; column++) {
                    table.Rows[0].Cells[column].Paragraphs[0].Text = headings[column];
                }
                for (int row = 1; row < 6; row++) {
                    table.Rows[row].Cells[0].Paragraphs[0].Text = $"ACC-{page:D3}-{row:D2}";
                    table.Rows[row].Cells[1].Paragraphs[0].Text = $"Owner {(page + row) % 17:D2}";
                    table.Rows[row].Cells[2].Paragraphs[0].Text = (page * 1000M + row * 37.25M).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
                    table.Rows[row].Cells[3].Paragraphs[0].Text = row % 3 == 0 ? "Review" : "Approved";
                }

                if (page == 1) {
                    WordChart chart = document.AddChart("Quarterly delivery", false, 400, 180);
                    chart.AddCategories(new List<string> { "Q1", "Q2", "Q3", "Q4" });
                    chart.AddBar("Completed", new List<int> { 12, 19, 26, 31 }, OfficeColor.Blue);
                    chart.AddBar("Pending", new List<int> { 6, 5, 4, 3 }, OfficeColor.Orange);
                } else if (page == 2) {
                    WordSmartArt smartArt = document.AddSmartArt(WordSmartArtType.BasicProcess);
                    smartArt.AddNode("Collect");
                    smartArt.AddNode("Validate");
                    smartArt.AddNode("Publish");
                } else if (page == 3) {
                    document.AddParagraph().AddImage(imagePath, 120, 60);
                    document.AddParagraph("Project link: ")
                        .AddHyperLink("OfficeIMO", new Uri("https://github.com/EvotecIT/OfficeIMO"), addStyle: true);
                }

                WordList list = document.AddListBulleted();
                list.AddItem($"Page {page:D3} validates list layout");
                list.AddItem("The fixture is reproducible and contains no private data");
                if (page < PageCount) {
                    document.AddPageBreak();
                }
            }

            document.Save();
            var pdfOptions = new PdfOptions();
            pdfOptions.RegisterFontFamily(
                PdfStandardFont.Helvetica,
                new PdfEmbeddedFontFamily("Carlito", PdfBenchmarkAssets.CarlitoRegular));
            PdfSaveResult saveResult = document.SaveAsPdf(pdfPath, new WordPdfSaveOptions {
                FontFamily = "Carlito",
                PdfOptions = pdfOptions,
                DefaultTableBorders = true
            });
            saveResult.RequireSuccess();
            string[] allowedDiagnosticCodes = {
                "NativeFontFamilySubstituted",
                "NativeBodyChartQuality",
                "NativeBodySmartArtUnsupported",
                "unsupported-font-ligature-substitution"
            };
            PdfConversionWarning[] unexpectedWarnings = saveResult.Warnings
                .Where(warning => warning.Severity != PdfConversionWarningSeverity.Information &&
                    !allowedDiagnosticCodes.Contains(warning.Code, StringComparer.Ordinal))
                .ToArray();
            if (unexpectedWarnings.Length > 0) {
                throw new InvalidOperationException(
                    "Rich Word PDF conversion reported an unexpected fidelity loss. " +
                    string.Join(" ", unexpectedWarnings.Select(static warning => warning.ToString())));
            }
            WriteConversionReport(conversionReportPath, saveResult);
        }

        using (WordDocument reopened = WordDocument.Load(docxPath)) {
            if (reopened.Tables.Count < PageCount || reopened.Charts.Count != 1 || reopened.SmartArts.Count != 1) {
                throw new InvalidDataException(
                    $"Rich Word fixture lost source structure: tables={reopened.Tables.Count}, " +
                    $"charts={reopened.Charts.Count}, smartArt={reopened.SmartArts.Count}.");
            }
        }

        ValidateGeneratedPdf(pdfPath);

        return new RichWordPdfCorpusArtifacts(docxPath, pdfPath, conversionReportPath);
    }

    private static void ValidateGeneratedPdf(string pdfPath) {
        byte[] bytes = File.ReadAllBytes(pdfPath);
        var readOptions = new PdfReadOptions { IncludeArtifactText = true };
        PdfReadDocument readDocument = PdfReadDocument.Open(bytes, readOptions);
        if (readDocument.Pages.Count != PageCount) {
            throw new InvalidDataException(
                $"Rich Word PDF has {readDocument.Pages.Count} pages; expected {PageCount}.");
        }

        string text = string.Join("\n", readDocument.Pages.Select(static page => page.ExtractText()));
        string[] requiredText = {
            "OfficeIMO rich interoperability corpus",
            "Generated benchmark fixture",
            "Quarterly delivery",
            "Q1",
            "Q4",
            "Project link",
            "OfficeIMO"
        };
        foreach (string marker in requiredText) {
            if (!text.Contains(marker, StringComparison.Ordinal)) {
                throw new InvalidDataException($"Rich Word PDF did not preserve required content '{marker}'.");
            }
        }

        OfficeIMO.Pdf.PdfDocument opened = OfficeIMO.Pdf.PdfDocument.Open(bytes, readOptions);
        if (opened.Read.Images().Count == 0) {
            throw new InvalidDataException("Rich Word PDF did not preserve its embedded image.");
        }

        const string expectedLink = "https://github.com/EvotecIT/OfficeIMO";
        PdfDocumentInfo info = opened.Inspect();
        if (!info.LinkUris.Contains(expectedLink, StringComparer.Ordinal)) {
            throw new InvalidDataException("Rich Word PDF did not preserve its external hyperlink annotation.");
        }
    }

    private static void WriteConversionReport(string path, PdfSaveResult result) {
        var payload = new {
            fidelityStatus = result.Report.FidelityStatus.ToString(),
            hasLoss = result.HasLoss,
            warnings = result.Warnings.Select(static warning => new {
                warning.Converter,
                warning.Code,
                warning.Source,
                warning.Message,
                severity = warning.Severity.ToString(),
                warning.Details
            })
        };
        File.WriteAllText(path, JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true }));
    }
}
