using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

// Content inside a placed paragraph frame keeps what the same paragraph has outside one.
public partial class Word {
    [Fact]
    public void SaveAsPdf_BookmarkInsideBodyFrameStaysALinkTarget() {
        using WordDocument document = CreateFrameDocument();
        WordParagraph framed = document.AddParagraph("FRAMEDTARGET");
        framed.AddBookmark("FrameTarget");
        AddFrameParagraph(framed, "around", "text", "margin", x: 1, width: 2880);
        document.AddParagraph("Body text beside the frame.");
        document.AddParagraph(string.Empty).AddHyperLink("GOTOFRAME", "FrameTarget");

        using PdfPigDocument pdf = RenderFramePdf(document);
        var page = pdf.GetPage(1);
        Assert.Contains("FRAMEDTARGET", page.Text);
        Assert.Contains(page.GetAnnotations(), annotation => annotation.Type == UglyToad.PdfPig.Annotations.AnnotationType.Link);
    }

    [Fact]
    public void SaveAsPdf_ListItemInsideFrameKeepsItsMarker() {
        using WordDocument document = CreateFrameDocument();
        WordList list = document.AddList(WordListStyle.LowerLetterWithDot);
        WordParagraph framed = list.AddItem("FRAMEDITEM");
        AddFrameParagraph(framed, "around", "text", "margin", x: 1, width: 2880);
        document.AddParagraph("Body text beside the frame.");

        using PdfPigDocument pdf = RenderFramePdf(document);
        var words = pdf.GetPage(1).GetWords().ToList();
        var item = words.Single(word => word.Text == "FRAMEDITEM");
        Assert.Contains(words, word => word.Text == "a." && Math.Abs(word.Letters[0].StartBaseLine.Y - item.Letters[0].StartBaseLine.Y) < .5D &&
            word.BoundingBox.Right <= item.BoundingBox.Left);
    }

    [Theory]
    [InlineData("header")]
    [InlineData("footer")]
    public void SaveAsPdf_PictureInsideRunningFrameUsesItsStoryImage(string story) {
        using WordDocument document = CreateFrameDocument();
        WordHeaderFooter target = story == "header" ? document.PageSetup.Header : document.PageSetup.Footer;
        WordParagraph framed = target.AddParagraph();
        framed.AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), OfficeIMO.OfficeLength.Pixels(32), OfficeIMO.OfficeLength.Pixels(32));
        AddFrameParagraph(framed, "around", "text", "margin", xAlign: "right", y: 1);
        target.AddParagraph("ANCHOR");
        document.AddParagraph("Body without pictures.");

        using PdfPigDocument pdf = RenderFramePdf(document);
        var image = Assert.Single(pdf.GetPage(1).GetImages());
        Assert.Equal(FramePageWidth - FrameMargin, image.Bounds.Right, 0);
    }

    [Theory]
    [InlineData("frame")]
    [InlineData("table")]
    public void SaveAsPdf_MarginAnchoredFloatInContinuousSectionUsesItsMargins(string kind) {
        // Word places a left-aligned, margin-anchored frame or floating table in a continuous
        // section against that section's left margin on the shared page.
        using WordDocument document = CreateFrameDocument();
        document.AddParagraph("Alpha");
        WordSection continuous = document.AddSection(WordSectionBreakType.Continuous);
        continuous.PageSetup.Margins.Left = OfficeIMO.OfficeLength.Points(144D);
        if (kind == "frame") {
            WordParagraph framed = continuous.AddParagraph("FLOATED");
            AddFrameParagraph(framed, "around", "text", "margin", xAlign: "left", y: 1, width: 2160);
        } else {
            WordTable table = continuous.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "FLOATED";
            table._tableProperties!.TablePositionProperties = new DocumentFormat.OpenXml.Wordprocessing.TablePositionProperties {
                VerticalAnchor = DocumentFormat.OpenXml.Wordprocessing.VerticalAnchorValues.Text,
                HorizontalAnchor = DocumentFormat.OpenXml.Wordprocessing.HorizontalAnchorValues.Margin,
                TablePositionXAlignment = DocumentFormat.OpenXml.Wordprocessing.HorizontalAlignmentValues.Left,
                TablePositionY = 1
            };
        }
        continuous.AddParagraph("Beside the float.");

        using PdfPigDocument pdf = RenderFramePdf(document);
        var floated = FindFrameWord(pdf, "FLOATED");
        Assert.InRange(floated.BoundingBox.Left, 143.5D, 150D);
    }
}
