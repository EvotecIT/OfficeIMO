using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.Excel;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfIndependentReverseConversionTests {
    // ReportLab supplies both inputs and labels. Full ordered text retains short
    // numbers, punctuation and repeated values that unique-token recall misses.
    [Theory]
    [InlineData("english-columns", "native")]
    [InlineData("english-columns", "docx")]
    [InlineData("english-columns", "html")]
    [InlineData("english-columns", "odt")]
    [InlineData("english-ledger", "native")]
    [InlineData("english-ledger", "docx")]
    [InlineData("english-ledger", "html")]
    [InlineData("english-ledger", "odt")]
    public void IndependentNativePdfRetainsCompleteOrderedText(string fixtureId, string route) {
        Fixture fixture = ReadFixture(fixtureId);
        PdfDocumentReadResult source = PdfDocumentReadResult.Load(fixture.Pdf);
        Assert.Single(source.Pages);

        string actual = route switch {
            "native" => source.Text,
            "docx" => ReadWordText(source),
            "html" => ReadHtmlText(source),
            "odt" => ReadOdtText(source),
            _ => throw new ArgumentOutOfRangeException(nameof(route))
        };

        // Only Unicode composition and whitespace are normalized. The full
        // sequence requires every occurrence, including the two 89.00 cells.
        Assert.Equal(Normalize(string.Join(" ", fixture.ReadingOrder)), Normalize(actual));
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("ods")]
    public void IndependentNativeLedgerRetainsEveryCellAndTypedAmount(string route) {
        Fixture fixture = ReadFixture("english-ledger");
        PdfDocumentReadResult source = PdfDocumentReadResult.Load(fixture.Pdf);
        PdfLogicalTable table = Assert.Single(source.Tables);
        Assert.Equal(fixture.Table.Length, table.Rows.Count);
        for (int row = 0; row < fixture.Table.Length; row++) Assert.Equal(fixture.Table[row], table.Rows[row]);

        // The independently authored labels establish the first row as headers.
        var options = new ExcelPdfImportOptions { UseFirstRowAsHeader = true };
        if (route == "xlsx") AssertExcelCells(source, fixture.Table, options);
        else AssertOdsCells(source, fixture.Table, options);
    }

    private static string ReadWordText(PdfDocumentReadResult source) {
        using OfficeIMO.Word.WordDocument document = (source.ConvertToWordWithEvidence()).Value;
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument reopened = WordprocessingDocument.Open(bytes, false);
        W.Body body = Assert.IsType<W.Body>(reopened.MainDocumentPart?.Document?.Body);
        return string.Join(" ", body.Descendants<W.Paragraph>()
            .Select(paragraph => string.Concat(paragraph.Descendants<W.Text>().Select(text => text.Text))));
    }

    private static string ReadHtmlText(PdfDocumentReadResult source) {
        string html = source.ToHtml(new PdfToHtmlOptions {
            Profile = PdfHtmlProfile.Semantic,
            IncludeMetadata = false,
            IncludeOutlines = false
        });
        using var document = new HtmlParser().ParseDocument(html);
        var page = Assert.Single(document.QuerySelectorAll(".pdf-page"));
        // Read saved HTML as DOM blocks; inserting whitespace around every tag
        // would split words whose styles produce adjacent inline fragments.
        return string.Join(" ", page.QuerySelectorAll("h1,h2,h3,h4,h5,h6,p,li,th,td")
            .Select(element => element.TextContent));
    }

    private static string ReadOdtText(PdfDocumentReadResult source) {
        byte[] artifact = source.ToOdtDocumentResult().Value.ToBytes();
        using var bytes = new MemoryStream(artifact);
        OdtDocument reopened = OdtDocument.Load(bytes);
        Assert.NotEmpty(reopened.Paragraphs);
        // The body-order paragraph view includes table-cell paragraphs. Its Text
        // property uses the canonical ODF codec for text:s, text:c, tabs and line
        // breaks; XElement.Value would silently drop those explicit primitives.
        return string.Join(" ", reopened.Paragraphs.Select(paragraph => paragraph.Text));
    }

    private static void AssertExcelCells(PdfDocumentReadResult source, string[][] expected, ExcelPdfImportOptions options) {
        using var bytes = new MemoryStream();
        var conversion = source.ConvertToExcelWithEvidence(options);
        using var convertedWorkbook = conversion.Value;
        convertedWorkbook.Save(bytes);
        PdfExcelTableImportReport report = conversion.Report;
        Assert.Equal(expected.Length - 1, Assert.Single(report.Entries).RowCount);
        Assert.True(report.HasOmittedPageContent); // The page title and closing prose are outside this table-only route.
        bytes.Position = 0;
        using SpreadsheetDocument reopened = SpreadsheetDocument.Open(bytes, false);
        var workbook = reopened.WorkbookPart!;
        S.SharedStringItem[] strings = workbook.SharedStringTablePart?.SharedStringTable?.Elements<S.SharedStringItem>().ToArray()
            ?? Array.Empty<S.SharedStringItem>();
        S.Row[] rows = Assert.Single(workbook.WorksheetParts).Worksheet!.Descendants<S.Row>().ToArray();
        Assert.Equal(expected.Length, rows.Length);
        for (int row = 0; row < expected.Length; row++) {
            S.Cell[] cells = rows[row].Elements<S.Cell>().ToArray();
            Assert.Equal(expected[row].Length, cells.Length);
            for (int column = 0; column < cells.Length; column++) {
                S.Cell cell = cells[column];
                Assert.Equal(((char)('A' + column)).ToString() + (row + 1).ToString(CultureInfo.InvariantCulture), cell.CellReference?.Value);
                if (row > 0 && column > 0) {
                    Assert.True(cell.DataType == null || cell.DataType.Value == S.CellValues.Number,
                        "Expected an editable numeric cell at " + cell.CellReference);
                    Assert.Equal(ParseNumber(expected[row][column]), ParseNumber(cell.CellValue!.Text));
                } else {
                    string actual = cell.DataType?.Value == S.CellValues.SharedString
                        ? strings[int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture)].InnerText
                        : cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? string.Empty;
                    Assert.Equal(expected[row][column], actual);
                }
            }
        }
    }

    private static void AssertOdsCells(PdfDocumentReadResult source, string[][] expected, ExcelPdfImportOptions options) {
        PdfOdsConversionResult result = source.ToOdsDocumentResult(options);
        Assert.Equal(expected.Length - 1, Assert.Single(result.Report.PdfReport.Entries).RowCount);
        Assert.True(result.Report.PdfReport.HasOmittedPageContent);
        using var bytes = new MemoryStream(result.Value.ToBytes());
        OdsDocument reopened = OdsDocument.Load(bytes);
        OdsSheet sheet = Assert.Single(reopened.Sheets);
        OdsUsedRange range = sheet.UsedRange!.Value;
        Assert.Equal(0L, range.FirstRow);
        Assert.Equal(0L, range.FirstColumn);
        Assert.Equal((long)expected.Length, range.RowCount);
        Assert.Equal((long)expected[0].Length, range.ColumnCount);
        for (int row = 0; row < expected.Length; row++) {
            for (int column = 0; column < expected[row].Length; column++) {
                OdsCell cell = sheet.Cell(row, column);
                if (row > 0 && column > 0) {
                    Assert.Equal(OdsCellValueKind.Number, cell.Value.Kind);
                    Assert.Equal(ParseNumber(expected[row][column]), cell.Value.AsDecimal());
                } else Assert.Equal(expected[row][column], cell.Text);
            }
        }
    }

    private static Fixture ReadFixture(string id) {
        string root = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "EnglishLayout");
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(root, "manifest.json")));
        JsonElement fixture = manifest.RootElement.GetProperty("cases").EnumerateArray()
            .Single(item => item.GetProperty("id").GetString() == id);
        byte[] pdf = File.ReadAllBytes(Path.Combine(root, id + "-native.pdf"));
        using SHA256 hash = SHA256.Create();
        string actualHash = string.Concat(hash.ComputeHash(pdf).Select(value => value.ToString("x2", CultureInfo.InvariantCulture)));
        Assert.Equal(fixture.GetProperty("files").GetProperty("native.pdf").GetString(), actualHash);
        return new Fixture(pdf,
            fixture.GetProperty("readingOrder").EnumerateArray().Select(value => value.GetString()!).ToArray(),
            fixture.GetProperty("table").EnumerateArray().Select(row => row.EnumerateArray().Select(value => value.GetString()!).ToArray()).ToArray());
    }

    private static decimal ParseNumber(string value) => decimal.Parse(value, NumberStyles.Number, CultureInfo.InvariantCulture);
    private static string Normalize(string value) => Regex.Replace(value.Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();
    private sealed record Fixture(byte[] Pdf, string[] ReadingOrder, string[][] Table);
}
