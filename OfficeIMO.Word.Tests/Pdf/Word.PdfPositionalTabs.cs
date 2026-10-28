using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", false)]
    [InlineData("footer", false)]
    [InlineData("table", false)]
    [InlineData("footer", true)]
    public void SaveAsPdf_PositionalTabsAlignCompleteMultiwordFields(string storyKind, bool splitFormatting) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph;
        if (storyKind == "table") {
            WordTable table = document.AddTable(1, 1);
            ConfigureNativeInputTable(table, 6400);
            paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        } else paragraph = storyKind == "footer" ? document.FooterDefaultOrCreate.AddParagraph() : document.AddParagraph();
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 10;
        paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(20D);
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
        paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        paragraph._run!.Append(new W.Text("Footer Left"), CreateNativeInputPositionalTab("center", false));
        if (splitFormatting) {
            paragraph._paragraph.Append(
                new W.Run((W.RunProperties)paragraph._run.RunProperties!.CloneNode(true), new W.Text("Footer ") { Space = SpaceProcessingModeValues.Preserve }),
                new W.Run(new W.RunProperties(new W.Bold(), new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "20" }), new W.Text("Middle")),
                new W.Run(CreateNativeInputPositionalTab("right", false)),
                new W.SimpleField(new W.Run((W.RunProperties)paragraph._run.RunProperties.CloneNode(true), new W.Text("Footer Right"))) { Instruction = " QUOTE \"Footer Right\" " });
        } else paragraph._run.Append(new W.Text("Footer Middle"), CreateNativeInputPositionalTab("right", false), new W.Text("Footer Right"));
        document.AddParagraph("BodyMarker");

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(NativeInputTableOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var footers = words.Where(word => word.Text == "Footer").OrderBy(word => word.BoundingBox.Left).ToArray();
        Assert.Equal(3, footers.Length);
        var left = Assert.Single(words, word => word.Text == "Left");
        var middle = Assert.Single(words, word => word.Text == "Middle");
        var right = Assert.Single(words, word => word.Text == "Right");
        Assert.InRange(footers[0].BoundingBox.Left, 59D, 61D);
        Assert.InRange((footers[1].BoundingBox.Left + middle.BoundingBox.Right) / 2D, 199D, 201D);
        Assert.InRange(right.BoundingBox.Right, 359D, 361D);
        foreach (var word in footers.Concat(new[] { middle, right }))
            Assert.Equal(left.Letters[0].StartBaseLine.Y, word.Letters[0].StartBaseLine.Y, 2);
    }

    [Theory]
    [InlineData("body", false, 0)]
    [InlineData("body", true, 0)]
    [InlineData("header", false, 0)]
    [InlineData("header", true, 0)]
    [InlineData("footer", false, 0)]
    [InlineData("footer", true, 0)]
    [InlineData("table", false, 0)]
    [InlineData("table", true, 0)]
    [InlineData("body", false, 30)]
    [InlineData("body", true, 30)]
    [InlineData("footer", false, 30)]
    [InlineData("footer", true, 30)]
    public void SaveAsPdf_PositionalTabsAlignRunsWithinMarginsOrParagraphIndents(string storyKind, bool relativeToIndent, int rightIndent) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph;
        if (storyKind == "table") {
            WordTable table = document.AddTable(1, 1);
            ConfigureNativeInputTable(table, 6400);
            paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        } else {
            paragraph = storyKind == "header" ? document.HeaderDefaultOrCreate.AddParagraph() :
                storyKind == "footer" ? document.FooterDefaultOrCreate.AddParagraph() : document.AddParagraph();
        }
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 10;
        paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(20D);
        paragraph.IndentationAfter = global::OfficeIMO.OfficeLength.Points(rightIndent);
        if (rightIndent > 0) {
            document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Append(new W.Style(
                new W.StyleName { Val = "Positional inherited indent" },
                new W.StyleParagraphProperties(new W.Indentation { Left = "400", Right = "600" })) {
                    Type = W.StyleValues.Paragraph, StyleId = "PositionalInheritedIndent"
                });
            paragraph._paragraph.ParagraphProperties!.Indentation = null;
            paragraph._paragraph.ParagraphProperties.ParagraphStyleId = new W.ParagraphStyleId { Val = "PositionalInheritedIndent" };
        }
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
        paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        paragraph._run!.Append(new W.Text("LEFT"),
            CreateNativeInputPositionalTab("center", relativeToIndent), new W.Text("CENTER"),
            CreateNativeInputPositionalTab("right", relativeToIndent), new W.Text("RIGHT"));
        document.AddParagraph("BodyMarker");
        Assert.Contains("LEFT\tCENTER\tRIGHT", paragraph.Text);
        string sourceXml = paragraph._paragraph.OuterXml;

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(NativeInputTableOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        var left = Assert.Single(page.GetWords(), word => word.Text == "LEFT");
        var center = Assert.Single(page.GetWords(), word => word.Text == "CENTER");
        var right = Assert.Single(page.GetWords(), word => word.Text == "RIGHT");
        Assert.InRange(left.BoundingBox.Left, 59D, 61D);
        double centerAnchor = relativeToIndent ? 210D : 200D;
        Assert.InRange((center.BoundingBox.Left + center.BoundingBox.Right) / 2D, centerAnchor - 1D, centerAnchor + 1D);
        double rightAnchor = 360D - rightIndent;
        Assert.InRange(right.BoundingBox.Right, rightAnchor - 1D, rightAnchor + 1D);
        Assert.Equal(left.BoundingBox.Bottom, center.BoundingBox.Bottom, 2);
        Assert.Equal(left.BoundingBox.Bottom, right.BoundingBox.Bottom, 2);
        Assert.Equal(sourceXml, paragraph._paragraph.OuterXml);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("footer")]
    [InlineData("table")]
    public void SaveAsPdf_HiddenPositionalTabsDoNotMoveFollowingVisibleText(string storyKind) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph;
        if (storyKind == "table") {
            WordTable table = document.AddTable(1, 1);
            ConfigureNativeInputTable(table, 6400);
            paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        } else paragraph = storyKind == "footer" ? document.FooterDefaultOrCreate.AddParagraph() : document.AddParagraph();
        paragraph.Text = "BEFORE";
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 10;
        paragraph._paragraph.Append(
            new W.Run(new W.RunProperties(new W.Vanish()), CreateNativeInputPositionalTab("right", false), new W.Text("HIDDEN")),
            new W.Run(new W.Text("AFTER")));
        document.AddParagraph("BodyMarker");
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(NativeInputTableOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        Assert.DoesNotContain("HIDDEN", page.Text);
        var before = Assert.Single(page.GetWords(), word => word.Text == "BEFORE");
        var after = Assert.Single(page.GetWords(), word => word.Text == "AFTER");
        Assert.InRange(after.BoundingBox.Left - before.BoundingBox.Right, -1D, 3D);
        Assert.Equal(before.Letters[0].StartBaseLine.Y, after.Letters[0].StartBaseLine.Y, 2);
        Assert.True(after.BoundingBox.Right < 160D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_RunningPositionalTabsAlignMaterializedPageFields(bool footer) {
        using WordDocument document = WordDocument.Create();
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordParagraph paragraph = story.AddParagraph();
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 10;
        paragraph._paragraph.Append(
            new W.Run(new W.Text("Section")),
            new W.Run(CreateNativeInputPositionalTab("right", false)),
            new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " PAGE " });
        document.AddParagraph("FirstBody");
        document.AddParagraph().AddBreak(WordBreakType.Page);
        document.AddParagraph("SecondBody");
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(NativeInputTableOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.NumberOfPages);
        foreach (int pageNumber in new[] { 1, 2 }) {
            var page = pdf.GetPage(pageNumber);
            Assert.DoesNotContain("999", page.Text);
            var number = Assert.Single(page.GetWords(), word => word.Text == pageNumber.ToString());
            Assert.InRange(number.BoundingBox.Right, 359D, 361D);
            Assert.Equal(Assert.Single(page.GetWords(), word => word.Text == "Section").BoundingBox.Bottom,
                number.BoundingBox.Bottom, 2);
        }
    }

    private static W.PositionalTab CreateNativeInputPositionalTab(string alignment, bool relativeToIndent) {
        W.PositionalTab tab = new();
        const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        tab.SetAttribute(new OpenXmlAttribute("w", "alignment", wordNamespace, alignment));
        tab.SetAttribute(new OpenXmlAttribute("w", "relativeTo", wordNamespace, relativeToIndent ? "indent" : "margin"));
        tab.SetAttribute(new OpenXmlAttribute("w", "leader", wordNamespace, "none"));
        return tab;
    }
}
