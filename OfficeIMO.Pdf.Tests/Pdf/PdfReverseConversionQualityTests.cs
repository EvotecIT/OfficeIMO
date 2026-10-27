using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Pdf.Word;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Regression tests for reconstructing editable documents from PDF pages.</summary>
public sealed class PdfReverseConversionQualityTests {
    private static string Fixture(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Rendering", "ReverseConversion", name);

    private static PdfDocument LoadFixture(string name) => PdfDocument.Load(File.ReadAllBytes(Fixture(name)));

    [Fact]
    public void TwoColumnArticleReadsEachColumnTopToBottom() {
        PdfDocumentReadResult result = LoadFixture("article-two-column.pdf").Read();
        string markdown = result.ExtractMarkdown();

        // Every paragraph keeps its own lines: no line of the neighbouring column is spliced in.
        Assert.Contains("Regional distribution networks have spent the last decade optimising for cost. The result is a set of long, thin supply chains", markdown, StringComparison.Ordinal);
        Assert.Contains("in a day. Teams that had not spent a week assembling the facts. Speed mattered more than precision", markdown, StringComparison.Ordinal);
        Assert.Contains("They had invested in shared data models with their suppliers and treated forecast accuracy", markdown, StringComparison.Ordinal);

        // The left column (with its figure caption and closing section) comes before the right column.
        int leftStart = markdown.IndexOf("Why resilience now", StringComparison.Ordinal);
        int leftEnd = markdown.IndexOf("The third difference is decision speed", StringComparison.Ordinal);
        int rightStart = markdown.IndexOf("in a day. Teams", StringComparison.Ordinal);
        int rightHeading = markdown.IndexOf("What it costs", StringComparison.Ordinal);
        Assert.True(leftStart >= 0 && leftEnd > leftStart, markdown);
        Assert.True(rightStart > leftEnd, markdown);
        Assert.True(rightHeading > rightStart, markdown);
        Assert.True(markdown.IndexOf("Building Resilient Supply Chains", StringComparison.Ordinal) < leftStart);
    }

    [Fact]
    public void TwoColumnArticleImportsInlineImagesAndSourceFontsInWord() {
        using var output = new MemoryStream();
        LoadFixture("article-two-column.pdf").SaveAsWord(output);
        using WordprocessingDocument document = WordprocessingDocument.Open(new MemoryStream(output.ToArray()), false);
        W.Body body = document.MainDocumentPart!.Document!.Body!;

        // Reflowed text cannot avoid a page-anchored image, so a figure with its own space is placed inline.
        Assert.NotEmpty(body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>());
        Assert.Empty(body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>());
        // Alt text comes from the figure caption rather than the PDF resource name.
        string? description = body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().First().Description?.Value;
        Assert.DoesNotContain("Im", description ?? string.Empty, StringComparison.Ordinal);

        // PostScript names are written as family names Word can resolve.
        string[] fonts = body.Descendants<W.RunFonts>().Select(fonts => fonts.Ascii?.Value ?? string.Empty).Distinct().ToArray();
        Assert.Contains("Liberation Serif", fonts);
        Assert.Contains("Liberation Sans", fonts);
        Assert.DoesNotContain(fonts, font => font.Contains('-') || font.Contains('+'));
    }

    [Fact]
    public void CommaDecimalAmountsAreReadAsDecimalsInExcel() {
        using var output = new MemoryStream();
        OfficeConversionReport report = LoadFixture("Pdf.Showcase.Statement.pdf").SaveAsExcel(output);
        double[] numbers = ReadNumericCells(output.ToArray());

        // "31,80 PLN" is 31.80, not 3180.
        Assert.Contains(31.8D, numbers);
        Assert.Contains(63.6D, numbers);
        Assert.DoesNotContain(3180D, numbers);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "pdf-excel-numeric-convention-mixed");
    }

    [Fact]
    public void InvoiceListPromotesItsHeaderAndTypesDatesAndAmounts() {
        using var output = new MemoryStream();
        LoadFixture("invoices.pdf").SaveAsExcel(output);
        using SpreadsheetDocument workbook = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        WorksheetPart sheet = workbook.WorkbookPart!.WorksheetParts.First();
        string[] header = ReadRowText(workbook, sheet, 1);
        Assert.Equal(new[] { "Invoice", "Customer", "Issued", "Due", "Amount" }, header);

        S.Row second = sheet.Worksheet!.Descendants<S.Row>().First(row => row.RowIndex?.Value == 2U);
        S.Cell[] cells = second.Elements<S.Cell>().ToArray();
        // ISO dates become date serials and "8 134,38" becomes 8134.38.
        Assert.Null(cells[2].DataType);
        Assert.True(double.Parse(cells[2].CellValue!.Text, CultureInfo.InvariantCulture) > 40000D);
        Assert.Equal(8134.38D, double.Parse(cells[4].CellValue!.Text, CultureInfo.InvariantCulture), 2);
    }

    [Fact]
    public void RightAlignedFiguresNextToTextColumnsStaySeparateColumns() {
        // Themed tables right-align numbers, so a figure ends close to the next left-aligned column.
        byte[] pdf = PdfDocument.Create()
            .H1("Quarterly results")
            .Paragraph(paragraph => paragraph.Text("Revenue improved in the current quarter."))
            .Table(new[] {
                new[] { "Region", "Revenue", "Active" },
                new[] { "North", "1250", "True" },
                new[] { "South", "980", "False" },
                new[] { "West", "1430", "True" }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(PdfDocument.Load(pdf).Read().Tables);
        Assert.Equal(3, table.Columns.Count);
        Assert.Equal(new[] { "Region", "Revenue", "Active" }, table.Rows[0]);
        Assert.Equal(new[] { "North", "1250", "True" }, table.Rows[1]);
    }

    [Theory]
    [InlineData("31,80", false, 0D)]
    [InlineData("1,234.50", true, 1234.5D)]
    [InlineData("1,234", true, 1234D)]
    [InlineData("12,34,567", false, 0D)]
    [InlineData("1.234,5", false, 0D)]
    public void InvariantNumericParsingRequiresValidDigitGrouping(string text, bool expected, double value) {
        bool parsed = PdfLogicalTableAnalysis.TryParseNumericValue(text, CultureInfo.InvariantCulture, out decimal result);
        Assert.Equal(expected, parsed);
        if (expected) Assert.Equal((decimal)value, result);
    }

    [Theory]
    [InlineData("31,80 PLN", "comma")]
    [InlineData("1.234,56", "comma")]
    [InlineData("1 234,56", "comma")]
    [InlineData("1,234,567", "point")]
    [InlineData("$4,210,000.00", "point")]
    [InlineData("1,234", "ambiguous")]
    [InlineData("23/12/2025", "none")]
    [InlineData("10.0.0.1", "none")]
    public void NumericSeparatorEvidenceIsClassifiedFromTheValue(string value, string expected) {
        var evidence = PdfExcelTableConverterExtensions.ClassifyNumericSeparators(value);
        string actual = evidence.Comma > 0 ? "comma" : evidence.Point > 0 ? "point" : evidence.Ambiguous != null ? "ambiguous" : "none";
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void DashboardTableKeepsHeaderFillWithItsWhiteText() {
        using var output = new MemoryStream();
        LoadFixture("Pdf.Showcase.Dashboard.pdf").SaveAsWord(output);
        using WordprocessingDocument document = WordprocessingDocument.Open(new MemoryStream(output.ToArray()), false);
        W.Body body = document.MainDocumentPart!.Document!.Body!;
        W.TableCell header = body.Descendants<W.TableCell>()
            .First(cell => cell.InnerText.Contains("Service line", StringComparison.Ordinal));
        string? fill = header.Descendants<W.Shading>().FirstOrDefault()?.Fill?.Value;
        Assert.False(string.IsNullOrEmpty(fill));
        Assert.NotEqual("FFFFFF", fill, StringComparer.OrdinalIgnoreCase);
        string? textColor = header.Descendants<W.Color>().FirstOrDefault()?.Val?.Value;
        Assert.Equal("FFFFFF", textColor, StringComparer.OrdinalIgnoreCase);
        Assert.NotEmpty(header.Descendants<W.Bold>());

        // KPI values set in a large font are values, not headings.
        Assert.DoesNotContain(body.Descendants<W.Paragraph>(), paragraph =>
            paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("Heading", StringComparison.Ordinal) == true &&
            !paragraph.InnerText.Any(char.IsLetter));
    }

    [Fact]
    public void VectorChartIsPlacedAsAPictureWithoutLooseLabels() {
        using var output = new MemoryStream();
        OfficeConversionReport report = LoadFixture("Pdf.Showcase.Dashboard.pdf").SaveAsWord(output);
        using WordprocessingDocument document = WordprocessingDocument.Open(new MemoryStream(output.ToArray()), false);
        W.Body body = document.MainDocumentPart!.Document!.Body!;

        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "pdf-vector-figure-rasterized");
        Assert.NotEmpty(body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>());
        // The axis and data labels are part of the picture, not separate paragraphs.
        Assert.DoesNotContain(body.Elements<W.Paragraph>(), paragraph => paragraph.InnerText.Trim() == "September");
        Assert.DoesNotContain(body.Elements<W.Paragraph>(), paragraph => paragraph.InnerText.Trim() == "812");
    }

    [Fact]
    public void DashboardMarkdownKeepsSignedPercentagesUnescaped() {
        string markdown = LoadFixture("Pdf.Showcase.Dashboard.pdf").Read().ExtractMarkdown();
        Assert.Contains("+2.4 pts", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("\\+", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("\\.", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void ArticleMarkdownWritesConsecutiveBulletsAsOneTightList() {
        string markdown = LoadFixture("article-two-column.pdf").Read().ExtractMarkdown().Replace("\r\n", "\n");
        Assert.Contains("- Shared inventory data with tier-one suppliers\n- Qualified alternatives for critical parts\n- Quarterly disruption rehearsals", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void WordPdfWordPdfRoundTripKeepsHeadingBookmarksResolvable() {
        byte[] firstPdf;
        using (WordDocument source = WordDocument.Create()) {
            source.AddTableOfContent();
            source.AddParagraph("Introduction").Style = WordParagraphStyles.Heading1;
            source.AddParagraph("Body text for the first section.");
            source.AddParagraph("Details").Style = WordParagraphStyles.Heading1;
            source.AddParagraph("Body text for the second section.");
            source.AddParagraph("More details").Style = WordParagraphStyles.Heading2;
            source.AddParagraph("Closing text.");
            using var stream = new MemoryStream();
            source.SaveAsPdf(stream);
            firstPdf = stream.ToArray();
        }

        using var docx = new MemoryStream();
        PdfDocument.Load(firstPdf).SaveAsWord(docx);
        using WordDocument reconstructed = WordDocument.Load(new MemoryStream(docx.ToArray()));
        using var secondPdf = new MemoryStream();
        // The second export used to fail on a bookmark that Word received but could not place.
        reconstructed.SaveAsPdf(secondPdf);
        Assert.True(secondPdf.Length > 0);

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(docx.ToArray()), false);
        foreach (W.Paragraph paragraph in package.MainDocumentPart!.Document!.Body!.Descendants<W.Paragraph>()) {
            Assert.True(paragraph.Elements<W.BookmarkStart>().Count() <= 1, paragraph.OuterXml);
        }
        string[] bookmarks = package.MainDocumentPart.Document.Body!.Descendants<W.BookmarkStart>().Select(b => b.Name!.Value!).ToArray();
        foreach (W.Hyperlink link in package.MainDocumentPart.Document.Body.Descendants<W.Hyperlink>().Where(link => link.Anchor != null)) {
            Assert.Contains(link.Anchor!.Value!, bookmarks);
        }
    }

    [Theory]
    [InlineData("ABCDEF+ArialMT", "Arial")]
    [InlineData("TimesNewRomanPS-BoldMT", "Times New Roman")]
    [InlineData("TimesNewRomanPSMT", "Times New Roman")]
    [InlineData("LiberationSans-Bold", "Liberation Sans")]
    [InlineData("BAAAAA+LiberationSerif", "Liberation Serif")]
    [InlineData("Arial,Bold", "Arial")]
    [InlineData("MS-Gothic", "MS Gothic")]
    [InlineData("CourierNewPSMT", "Courier New")]
    [InlineData("SourceSansPro-Regular", "Source Sans Pro")]
    [InlineData("Calibri", "Calibri")]
    public void PostScriptFontNamesMapToFamilyNames(string baseFont, string expected) {
        Assert.Equal(expected, PdfFontFamilyNames.ToFamilyName(baseFont));
    }

    [Theory]
    [InlineData("1 / 2", true)]
    [InlineData("Page 3 of 12", true)]
    [InlineData("Chapter 1", false)]
    [InlineData("- 4 -", true)]
    [InlineData("3 / 12", true)]
    [InlineData("Revenue grew 12%", false)]
    [InlineData("2026 annual report", false)]
    public void PageNumberTextIsRecognizedWithoutVocabulary(string text, bool expected) {
        Assert.Equal(expected, PdfDocumentSemanticEnricher.IsPageNumberText(text));
    }

    [Fact]
    public void PowerPointReconstructionUsesNeutralTableStyleAndUnwrappedSourceLines() {
        using var output = new MemoryStream();
        LoadFixture("invoices.pdf").SaveAsPowerPoint(output);
        using PresentationDocument presentation = PresentationDocument.Open(new MemoryStream(output.ToArray()), false);
        DocumentFormat.OpenXml.Presentation.Slide[] slides = presentation.PresentationPart!.SlideParts.Select(static part => part.Slide!).ToArray();
        string[] styleIds = slides.SelectMany(static slide => slide.Descendants<A.TableStyleId>()).Select(static id => id.Text).ToArray();
        Assert.NotEmpty(styleIds);
        Assert.All(styleIds, id => Assert.Equal("{5940675A-B579-460E-94D1-54222C63F5DA}", id));
        // Source lines never wrap, and font names are families PowerPoint resolves.
        DocumentFormat.OpenXml.Presentation.Shape[] textBoxes = slides.SelectMany(static slide => slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>())
            .Where(shape => shape.TextBody != null && shape.TextBody.InnerText.Length > 0).ToArray();
        Assert.NotEmpty(textBoxes);
        Assert.All(textBoxes, shape => Assert.Equal(A.TextWrappingValues.None, shape.TextBody!.BodyProperties!.Wrap!.Value));
        Assert.DoesNotContain(slides.SelectMany(static slide => slide.Descendants<A.LatinFont>()), font => font.Typeface?.Value?.Contains('-') == true);
    }

    private static double[] ReadNumericCells(byte[] xlsx) {
        using SpreadsheetDocument workbook = SpreadsheetDocument.Open(new MemoryStream(xlsx), false);
        return workbook.WorkbookPart!.WorksheetParts
            .SelectMany(part => part.Worksheet!.Descendants<S.Cell>())
            .Where(cell => cell.DataType == null && cell.CellValue != null)
            .Select(cell => double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture))
            .ToArray();
    }

    private static string[] ReadRowText(SpreadsheetDocument workbook, WorksheetPart sheet, uint rowIndex) {
        S.SharedStringTable? shared = workbook.WorkbookPart!.SharedStringTablePart?.SharedStringTable;
        S.Row row = sheet.Worksheet!.Descendants<S.Row>().First(candidate => candidate.RowIndex?.Value == rowIndex);
        return row.Elements<S.Cell>().Select(cell => {
            if (cell.DataType?.Value == S.CellValues.SharedString && shared != null) {
                return shared.ElementAt(int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture)).InnerText;
            }
            if (cell.DataType?.Value == S.CellValues.InlineString) return cell.InlineString?.InnerText ?? string.Empty;
            return cell.CellValue?.Text ?? string.Empty;
        }).ToArray();
    }
}
