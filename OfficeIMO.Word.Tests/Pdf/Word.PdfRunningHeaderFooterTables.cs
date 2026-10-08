using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    private static WordTable CreateRunningStoryTable(WordHeaderFooter story, string text, double height, WordTextDirection direction) {
        WordTable table = story.AddTable(1, 1, WordTableStyle.TableGrid);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa; table.Width = 2400;
        table.GridColumnWidth = new List<int> { 2400 };
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.WidthType = WordTableWidthUnit.Dxa; cell.Width = 2400; cell.TextDirection = direction;
        table.Rows[0].Height = (int)(height * 20);
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = text; paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
        story.AddParagraph();
        return table;
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void RunningHeaderFooterTablePreservesBordersDirectionAndDocumentSource(bool footer, bool up, bool nativeDoc) {
        using WordDocument source = CreateJoinedParagraphDocument();
        source.AddParagraph("BODY");
        source.Sections[0].Margins.HeaderDistance = 360;
        source.Sections[0].Margins.FooterDistance = 360;
        WordHeaderFooter story = footer ? source.FooterDefaultOrCreate : source.HeaderDefaultOrCreate;
        CreateRunningStoryTable(story, "HEADER", 120,
            up ? WordTextDirection.BottomToTopLeftToRight : WordTextDirection.TopToBottomRightToLeft);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        Assert.Empty(document.ValidateDocument());
        string StoryXml() => document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml +
            string.Concat(document._wordprocessingDocument.MainDocumentPart.HeaderParts.Select(part => part.Header.OuterXml)) +
            string.Concat(document._wordprocessingDocument.MainDocumentPart.FooterParts.Select(part => part.Footer.OuterXml));
        string before = StoryXml();
        using var pdf = OpenJoinedParagraphPdf(document);
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("HEADER", page.Text); Assert.Contains("BODY", page.Text);
        var letter = page.Letters.First(item => item.Value == "H");
        double vertical = letter.EndBaseLine.Y - letter.StartBaseLine.Y;
        Assert.True(up ? vertical > 1 : vertical < -1);
        var edges = page.Paths.Where(path => path.IsStroked)
            .Select(path => path.GetBoundingRectangle()).Where(bounds => bounds.HasValue).Select(bounds => bounds!.Value).ToArray();
        Assert.NotEmpty(edges);
        double top = edges.Max(edge => edge.Top), bottom = edges.Min(edge => edge.Bottom);
        Assert.InRange(top - bottom, 119.99, 120.01);
        if (!footer) {
            Assert.InRange(top, page.Height - 18.01, page.Height - 17.99);
            Assert.True(page.Letters.First(item => item.Value == "B").BoundingBox.Top < bottom);
        }
        Assert.Equal(before, StoryXml());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RunningHeaderFooterVariantsCalculateFieldsInsideTables(bool footer) {
        using WordDocument document = CreateJoinedParagraphDocument();
        document.Sections[0].Margins.HeaderDistance = 360;
        document.Sections[0].Margins.FooterDistance = 360;
        WordHeaderFooter[] stories = footer
            ? new WordHeaderFooter[] { document.FooterFirstOrCreate, document.FooterEvenOrCreate, document.FooterDefaultOrCreate }
            : new WordHeaderFooter[] { document.HeaderFirstOrCreate, document.HeaderEvenOrCreate, document.HeaderDefaultOrCreate };
        string[] labels = { "First", "Even", "Default" };
        double[] heights = { 24, 72, 120 };
        for (int index = 0; index < stories.Length; index++) {
            WordParagraph paragraph = CreateRunningStoryTable(stories[index], labels[index], heights[index],
                WordTextDirection.TopToBottomRightToLeft).Rows[0].Cells[0].Paragraphs[0];
            paragraph._paragraph.Append(new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " PAGE " },
                new W.Run(new W.Text("/")), new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " NUMPAGES " });
        }
        document.AddParagraph("BODY"); document.AddParagraph().AddBreak(WordBreakType.Page);
        document.AddParagraph("BODY"); document.AddParagraph().AddBreak(WordBreakType.Page); document.AddParagraph("BODY");
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Equal(3, pdf.NumberOfPages);
        for (int index = 0; index < 3; index++) {
            string text = string.Concat(pdf.GetPage(index + 1).Letters.Select(letter => letter.Value));
            Assert.Contains(labels[index] + (index + 1) + "/3", text);
            Assert.DoesNotContain("999", text);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RunningHeaderFooterUncachedFieldsKeepIndividualNumberStylesAndLiteralBraces(bool footer, bool complex) {
        using WordDocument document = CreateJoinedParagraphDocument(); document.AddParagraph("BODY");
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        CreateRunningStoryTable(story, "TABLE", 24, WordTextDirection.LeftToRightTopToBottom);
        WordParagraph paragraph = story.AddParagraph("VISIBLE {page} ");
        if (complex) paragraph._paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" PAGE \\* roman ")), new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        else paragraph._paragraph.Append(new W.SimpleField { Instruction = " PAGE \\* roman " });
        paragraph._paragraph.Append(new W.Run(new W.Text("/")), new W.SimpleField { Instruction = " NUMPAGES \\* Arabic " });
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Contains("VISIBLE {page} i/1", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RunningHeaderFooterJoinedParagraphFieldsUseCurrentPageValues(bool footer) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        CreateRunningStoryTable(story, "TABLE", 24, WordTextDirection.LeftToRightTopToBottom);
        WordParagraph first = story.AddParagraph("JOIN");
        first._paragraph.Append(new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " PAGE " });
        HideJoinMark(first, true);
        story.AddParagraph("END");
        document.AddParagraph("BODY"); document.AddParagraph().AddBreak(WordBreakType.Page); document.AddParagraph("BODY");
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Equal(2, pdf.NumberOfPages);
        for (int index = 1; index <= 2; index++) {
            Assert.Contains($"JOIN{index}END", pdf.GetPage(index).Text);
            Assert.DoesNotContain("999", pdf.GetPage(index).Text);
        }
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void RunningHeaderFooterSectionCountsAreIndependentOfVisibleNumbering(bool footer, bool table, bool continuingSections) {
        using WordDocument document = CreateJoinedParagraphDocument();
        void AddStory(int sectionIndex, string label) {
            WordHeaderFooter story = footer ? RequireSectionFooter(document, sectionIndex, W.HeaderFooterValues.Default)
                : RequireSectionHeader(document, sectionIndex, W.HeaderFooterValues.Default);
            WordParagraph paragraph = table
                ? CreateRunningStoryTable(story, label, 24, WordTextDirection.LeftToRightTopToBottom).Rows[0].Cells[0].Paragraphs[0]
                : story.AddParagraph(label);
            foreach (string field in new[] { "PAGE", "SECTIONPAGES", "NUMPAGES" }) {
                paragraph._paragraph.Append(new W.Run(new W.Text("/")),
                    new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " " + field + " " });
            }
        }
        if (!continuingSections) document.Sections[0].AddPageNumbering(5, WordNumberFormat.Decimal);
        AddStory(0, "First");
        document.AddParagraph("BODY"); document.AddParagraph().AddBreak(WordBreakType.Page); document.AddParagraph("BODY");
        if (continuingSections) {
            WordSection second = document.AddSection();
            AddStory(1, "Second");
            second.AddParagraph("BODY"); document.AddParagraph().AddBreak(WordBreakType.Page); second.AddParagraph("BODY");
        }
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Equal(continuingSections ? 4 : 2, pdf.NumberOfPages);
        for (int page = 1; page <= pdf.NumberOfPages; page++) {
            string label = page <= 2 ? "First" : "Second";
            int visible = continuingSections ? page : page + 4;
            Assert.Contains($"{label}/{visible}/2/{pdf.NumberOfPages}", pdf.GetPage(page).Text);
            Assert.DoesNotContain("999", pdf.GetPage(page).Text);
        }
    }
}
