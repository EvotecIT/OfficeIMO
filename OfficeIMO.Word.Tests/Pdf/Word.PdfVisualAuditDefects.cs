using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    private static PdfPigDocument OpenAuditPdf(WordDocument document, WordPdfExportOptions? options = null) =>
        PdfPigDocument.Open(document.ToPdf(options).ToBytes(new PdfSaveOptions { AllowLoss = true }));

    [Fact]
    public void SaveAsPdf_FooterPageNumberInsideBlockContentControlIsRendered() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("First page");
        document.AddPageBreak();
        document.AddParagraph("Second page");
        document.AddHeadersAndFooters();
        document.Footer!.Default.AddPageNumber(WordPageNumberStyle.PlainNumber);

        Assert.Contains(document.Footer.Default.Elements, element => element is WordStructuredDocumentTag);

        using PdfPigDocument pdf = OpenAuditPdf(document);
        Assert.Equal(2, pdf.NumberOfPages);
        for (int pageNumber = 1; pageNumber <= 2; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            var footerWords = page.GetWords().Where(word => word.BoundingBox.Bottom < page.Height / 4D).ToList();
            Assert.Contains(footerWords, word => word.Text == pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }

    private static WordDocument CreateTrackedChangeDocument() {
        WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Clause: ");
        Paragraph source = paragraph._paragraph;
        source.Append(new InsertedRun(new Run(new Text("added words") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })) {
            Id = "1", Author = "Reviewer", Date = new DateTime(2026, 10, 1)
        });
        source.Append(new DeletedRun(new Run(new DeletedText(" removed words") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })) {
            Id = "2", Author = "Reviewer", Date = new DateTime(2026, 10, 1)
        });
        return document;
    }

    [Fact]
    public void SaveAsPdf_ExplicitFinalViewRendersCurrentTextAndReportsOmittedRevisions() {
        using WordDocument document = CreateTrackedChangeDocument();
        var options = new WordPdfExportOptions { TrackedChanges = WordPdfTrackedChangeView.Final };
        using (document.ToPdf(options, out OfficeConversionReport report))
            Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "tracked-revisions-projected");

        using PdfPigDocument pdf = OpenAuditPdf(document, options);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("added words", text);
        Assert.DoesNotContain("removed words", text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_DefaultAndExplicitMarkupShowRevisionsWithoutFinalProjectionWarning(bool explicitMarkup) {
        using WordDocument document = CreateTrackedChangeDocument();
        WordPdfExportOptions? options = explicitMarkup
            ? new WordPdfExportOptions { TrackedChanges = WordPdfTrackedChangeView.Markup } : null;
        using (document.ToPdf(options, out OfficeConversionReport report))
            Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "tracked-revisions-projected");
        using PdfPigDocument pdf = OpenAuditPdf(document, options);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("added words", text);
        Assert.Contains("removed words", text);
    }

    [Fact]
    public void SaveAsPdf_DocumentWithoutTrackedChangesHasNoRevisionDiagnostic() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Plain text");
        using (document.ToPdf(null, out OfficeConversionReport report))
            Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "tracked-revisions-projected");
    }

    [Theory]
    [InlineData(WordListStyle.BulletedChars)]
    [InlineData(WordListStyle.Bulleted)]
    public void SaveAsPdf_SymbolFontBulletsRenderAsUnicodeGlyphs(WordListStyle listStyle) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(listStyle);
        list.AddItem("Level 0", 0);
        list.AddItem("Level 1", 1);
        list.AddItem("Level 2", 2);

        using PdfPigDocument pdf = OpenAuditPdf(document);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Level 2", text);
        // Symbol-font codes such as "v" or "§" must never print as letters.
        Assert.DoesNotContain("§", text);
        Assert.DoesNotContain("vLevel", text.Replace(" ", string.Empty));
    }

    [Theory]
    [InlineData("§", "Wingdings", "▪")]
    [InlineData("v", "Wingdings", "❖")]
    [InlineData("", "Wingdings", "▪")]
    [InlineData("·", "Symbol", "•")]
    [InlineData("", "Symbol", "•")]
    [InlineData("¨", "Symbol", "♦")]
    [InlineData("o", "Courier New", "o")]
    [InlineData("\u0179", "Symbol", "\u0179")]
    public void ListMarkerNormalizer_MapsSymbolFontCodesToUnicode(string marker, string font, string expected) {
        (string normalized, _) = OfficeIMO.Drawing.OfficeTextListMarkerNormalizer.Normalize(marker, font);
        Assert.Equal(expected, normalized);
    }

    [Fact]
    public void SaveAsPdf_NarrowTableKeepsCellMarginsInsteadOfFailing() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(3, 6);
        table.WidthType = WordTableWidthUnit.Pct;
        table.Width = 100; // 100 fiftieths of a percent: a 2% wide table.
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Longer text example to show wrapping";

        using PdfPigDocument pdf = OpenAuditPdf(document);
        Assert.Contains("wrapping", pdf.GetPage(1).Text);
    }

    [Fact]
    public void ToPdf_RunningStoriesReportDisposedSourceDocument() {
        PdfDocument pdf;
        using (WordDocument document = WordDocument.Create()) {
            document.AddParagraph("Body");
            document.AddHeadersAndFooters();
            document.Footer!.Default.AddPageNumber(WordPageNumberStyle.PlainNumber);
            pdf = document.ToPdf();
        }

        Exception error = Assert.ThrowsAny<Exception>(() => pdf.ToBytes(new PdfSaveOptions { AllowLoss = true }));
        Assert.True(error is ObjectDisposedException || error.InnerException is ObjectDisposedException, error.ToString());
    }

    [Fact]
    public void SaveAsPdf_RepeatingSectionItemsRenderAndUseWordItemControls() {
        using WordDocument document = WordDocument.Create();
        WordRepeatingSection section = document.AddParagraph("Tasks:").AddRepeatingSection("Tasks", "Tasks", "Tasks");
        section.SetTextItems(new[] { "Collect source", "Generate document", "Validate output" });

        Assert.Equal(new[] { "Collect source", "Generate document", "Validate output" }, section.TextItems);
        var items = document._wordprocessingDocument.MainDocumentPart!.Document.Body!.Descendants<SdtRun>()
            .Where(sdt => sdt.SdtProperties?.Elements<DocumentFormat.OpenXml.Office2013.Word.SdtRepeatedSectionItem>().Any() == true)
            .ToList();
        Assert.Equal(3, items.Count);

        using PdfPigDocument pdf = OpenAuditPdf(document);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Collect source", text);
        Assert.Contains("Validate output", text);
    }

    [Fact]
    public void SaveAsPdf_NormalTableFirstRowIsNotBold() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2, WordTableStyle.TableNormal);
        table.ConditionalFormattingFirstRow = true;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Address";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "Street";

        using PdfPigDocument pdf = OpenAuditPdf(document);
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "A").ToList();
        Assert.NotEmpty(letters);
        Assert.DoesNotContain(letters, letter => letter.FontName.Contains("Bold", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("start", WordListLevelAlignment.Left)]
    [InlineData("end", WordListLevelAlignment.Right)]
    [InlineData("center", WordListLevelAlignment.Center)]
    [InlineData("both", WordListLevelAlignment.Left)]
    public void SaveAsPdf_AcceptsEveryListLevelJustificationToken(string token, WordListLevelAlignment expected) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Bulleted);
        list.AddItem("Item");
        var level = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Descendants<Level>().First();
        level.LevelJustification!.Val!.InnerText = token;
        foreach (var justification in document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                     .Descendants<LevelJustification>())
            justification.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("w", "val",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", token));

        Assert.Equal(expected, list.Numbering.Levels[0].LevelJustification);
        using PdfPigDocument pdf = OpenAuditPdf(document);
        Assert.Contains("Item", pdf.GetPage(1).Text);
    }
}

public partial class Word {
    [Fact]
    public void SaveAsPdf_DisplayMathIsCentredLikeWord() {
        using WordDocument document = WordDocument.Create();
        document.AddEquation("<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara>");

        using PdfPigDocument pdf = OpenAuditPdf(document);
        var page = pdf.GetPage(1);
        var letters = page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToList();
        Assert.NotEmpty(letters);
        double centre = (letters.Min(letter => letter.GlyphRectangle.Left) + letters.Max(letter => letter.GlyphRectangle.Right)) / 2D;
        Assert.InRange(centre, page.Width / 2D - 20D, page.Width / 2D + 20D);
    }
}
