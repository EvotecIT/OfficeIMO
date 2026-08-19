using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReverseConversionScorecardTests {
    [Fact]
    public void Scorecard_ExecutesTwentyEightCrossProducerReverseConversions() {
        string repositoryRoot = FindRepositoryRoot();
        using JsonDocument scorecard = JsonDocument.Parse(File.ReadAllBytes(
            Path.Combine(repositoryRoot, "Docs", "pdf-reverse-conversion-scorecard.json")));
        JsonElement root = scorecard.RootElement;
        JsonElement[] producers = root.GetProperty("producers").EnumerateArray().ToArray();
        JsonElement[] routes = root.GetProperty("routes").EnumerateArray().ToArray();
        string[] routeIds = routes
            .Select(static route => route.GetProperty("id").GetString()!)
            .ToArray();

        Assert.Equal(7, producers.Length);
        Assert.Equal(new[] { "pdf-to-docx", "pdf-to-html", "pdf-to-xlsx", "pdf-to-pptx" }, routeIds);
        Assert.Equal(producers.Length * routeIds.Length, root.GetProperty("matrixCases").GetInt32());

        var executed = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement producer in producers) {
            string producerId = producer.GetProperty("id").GetString()!;
            string sourcePath = Path.Combine(repositoryRoot, producer.GetProperty("path").GetString()!.Replace('/', Path.DirectorySeparatorChar));
            Assert.True(File.Exists(sourcePath), "Missing scorecard fixture: " + sourcePath);
            byte[] source = File.ReadAllBytes(sourcePath);
            PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(source);
            Assert.NotEmpty(logical.Pages);
            HashSet<string> sourceTokens = GetTokens(string.Join(" ", logical.Pages.SelectMany(static page => page.TextBlocks).Select(static block => block.Text)));
            Assert.NotEmpty(sourceTokens);
            bool expectedTables = producer.GetProperty("expectedTables").GetBoolean();
            if (expectedTables) Assert.NotEmpty(logical.Tables);

            foreach (JsonElement routeConfiguration in routes) {
                string route = routeConfiguration.GetProperty("id").GetString()!;
                Assert.True(executed.Add(producerId + ":" + route));
                ExecuteAndReopen(routeConfiguration, source, logical, sourceTokens, expectedTables);
            }
        }

        Assert.Equal(28, executed.Count);
    }

    private static void ExecuteAndReopen(
        JsonElement routeConfiguration,
        byte[] source,
        PdfCore.PdfLogicalDocument logical,
        HashSet<string> sourceTokens,
        bool expectedTables) {
        string route = routeConfiguration.GetProperty("id").GetString()!;
        switch (route) {
            case "pdf-to-docx":
                using (OfficeIMO.Word.WordDocument document = logical.ToWordDocument()) {
                    using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(document.ToBytes()), false);
                    DocumentFormat.OpenXml.Wordprocessing.Body body = Assert.IsType<DocumentFormat.OpenXml.Wordprocessing.Body>(package.MainDocumentPart?.Document?.Body);
                    AssertTokenRecall(sourceTokens, string.Join(" ", body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(static text => text.Text)), routeConfiguration.GetProperty("minimumTokenRecall").GetDouble(), route);
                    if (expectedTables) Assert.NotEmpty(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>());
                }
                return;
            case "pdf-to-html":
                string html = logical.ToHtml(new PdfHtmlSaveOptions { Profile = PdfHtmlProfile.Semantic });
                Assert.Contains("<!doctype html", html, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(logical.Pages.Count, Regex.Matches(html, "class=\"pdf-page\"[^>]*data-page-number=", RegexOptions.CultureInvariant).Count);
                AssertTokenRecall(sourceTokens, Regex.Replace(html, "<[^>]+>", " "), routeConfiguration.GetProperty("minimumTokenRecall").GetDouble(), route);
                return;
            case "pdf-to-xlsx":
                using (var stream = new MemoryStream()) {
                    PdfExcelTableImportReport report = logical.SaveTablesAsExcel(stream);
                    byte[] workbook = stream.ToArray();
                    if (expectedTables) Assert.NotEmpty(report.Entries);
                    HashSet<string>? tableTokens = null;
                    if (report.Entries.Count > 0) {
                        tableTokens = GetTokens(string.Join(" ", logical.Tables.SelectMany(static table => {
                            PdfCore.PdfLogicalTableData data = PdfCore.PdfLogicalTableAnalysis.Extract(table);
                            return data.Columns.Concat(data.Rows.SelectMany(static row => row));
                        })));
                    } else {
                        Assert.True(report.HasOmittedPageContent, "A table-only XLSX projection must report omitted non-table page content.");
                    }
                    using SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(workbook), false);
                    Assert.NotNull(package.WorkbookPart?.Workbook);
                    if (report.Entries.Count > 0) {
                        Assert.NotEmpty(package.WorkbookPart!.WorksheetParts.SelectMany(static worksheet => worksheet.TableDefinitionParts));
                        AssertTokenRecall(tableTokens!, GetSpreadsheetText(package), routeConfiguration.GetProperty("minimumTableTokenRecall").GetDouble(), route);
                    }
                }
                return;
            case "pdf-to-pptx":
                PdfPowerPointConversionResult result = PdfCore.PdfDocument.Open(source)
                    .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateEditableContent());
                using (result.Value) {
                    using var stream = new MemoryStream();
                    result.Value.Save(stream);
                    using PresentationDocument package = PresentationDocument.Open(new MemoryStream(stream.ToArray()), false);
                    Assert.Equal(logical.Pages.Count, package.PresentationPart!.SlideParts.Count());
                    string presentationText = string.Join(" ", package.PresentationPart.SlideParts.SelectMany(static slide => slide.Slide.Descendants<A.Text>()).Select(static text => text.Text));
                    AssertTokenRecall(sourceTokens, presentationText, routeConfiguration.GetProperty("minimumTokenRecall").GetDouble(), route);
                    if (expectedTables) Assert.NotEmpty(package.PresentationPart.SlideParts.SelectMany(static slide => slide.Slide.Descendants<A.Table>()));
                }
                return;
            default:
                throw new InvalidOperationException("Unknown reverse-conversion route: " + route);
        }
    }

    private static HashSet<string> GetTokens(string value) => Regex.Matches(value.ToLowerInvariant(), @"[\p{L}\p{N}]{4,}", RegexOptions.CultureInvariant)
        .Cast<Match>()
        .Select(static match => match.Value)
        .ToHashSet(StringComparer.Ordinal);

    private static void AssertTokenRecall(HashSet<string> expectedTokens, string actualText, double minimumRecall, string route) {
        Assert.NotEmpty(expectedTokens);
        HashSet<string> actualTokens = GetTokens(actualText);
        int retained = expectedTokens.Count(actualTokens.Contains);
        double recall = (double)retained / expectedTokens.Count;
        Assert.True(recall >= minimumRecall, route + " retained " + retained + "/" + expectedTokens.Count + " source tokens (" + recall.ToString("P1", System.Globalization.CultureInfo.InvariantCulture) + "); expected at least " + minimumRecall.ToString("P1", System.Globalization.CultureInfo.InvariantCulture) + ".");
    }

    private static string GetSpreadsheetText(SpreadsheetDocument package) {
        string[] sharedStrings = package.WorkbookPart?.SharedStringTablePart?.SharedStringTable?
            .Elements<S.SharedStringItem>()
            .Select(static item => item.InnerText)
            .ToArray() ?? Array.Empty<string>();
        var values = new List<string>();
        foreach (S.Cell cell in package.WorkbookPart!.WorksheetParts.SelectMany(static worksheet => worksheet.Worksheet.Descendants<S.Cell>())) {
            if (cell.DataType?.Value == S.CellValues.SharedString &&
                int.TryParse(cell.CellValue?.Text, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int sharedIndex) &&
                sharedIndex >= 0 && sharedIndex < sharedStrings.Length) {
                values.Add(sharedStrings[sharedIndex]);
            } else if (cell.InlineString is not null) {
                values.Add(cell.InlineString.InnerText);
            } else if (cell.CellValue is not null) {
                values.Add(cell.CellValue.Text);
            }
        }
        return string.Join(" ", values);
    }

    private static string FindRepositoryRoot() {
        string? current = AppContext.BaseDirectory;
        while (!string.IsNullOrWhiteSpace(current)) {
            if (File.Exists(Path.Combine(current, "OfficeIMO.sln"))) return current;
            current = Directory.GetParent(current)?.FullName;
        }
        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }
}
