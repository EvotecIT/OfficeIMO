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

    [Fact]
    public void SaveAsPdf_FramedParagraphWithPageBreakBeforeStartsANewPage() {
        using WordDocument document = CreateFrameDocument();
        document.AddParagraph("Intro.");
        WordParagraph framed = document.AddParagraph("FRAMED");
        framed.PageBreakBeforeOverride = true;
        AddFrameParagraph(framed, "notBeside", "text", "margin", x: 1, width: 2880);
        using PdfPigDocument pdf = RenderFramePdf(document);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("FRAMED", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData("shading")]
    [InlineData("border")]
    [InlineData("chart")]
    public void SaveAsPdf_DecoratedFrameKeepsOrdinaryParagraphRendering(string content) {
        // Shading, borders and charts are drawn by the ordinary paragraph route, so such a frame
        // stays on it and reports the approximation instead of losing them.
        using WordDocument document = CreateFrameDocument();
        WordParagraph framed;
        if (content == "chart") {
            WordChart chart = document.AddChart("Frame chart");
            chart.AddCategories(new[] { "A", "B" }.ToList());
            chart.AddBar("Values", new[] { 2, 3 }, OfficeIMO.OfficeColor.ParseHex("276749"));
            framed = document.Paragraphs.Last(paragraph => paragraph.Chart != null);
        } else {
            framed = document.AddParagraph("FRAMED");
            if (content == "shading") framed.ShadingFillColor = OfficeIMO.OfficeColor.Parse("E6F2FF");
            else {
                framed.Borders.BottomStyle = WordBorderStyle.Single;
                framed.Borders.BottomColor = OfficeIMO.OfficeColor.Parse("FF0000");
                framed.Borders.BottomSize = OfficeIMO.OfficeLength.Points(2);
            }
        }
        AddFrameParagraph(framed, "around", "text", "margin", x: 1, width: 4320);
        document.AddParagraph("Body text.");

        var result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.Contains(result.Warnings, warning => warning.Code == "native-paragraph-frame-flattened");
    }

    [Fact]
    public void SaveAsPdf_DefaultTableStyleDoesNotFormatFrameText() {
        using WordDocument document = CreateFrameDocument();
        var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (var existing in styles.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>()
                     .Where(style => style.Type?.Value == DocumentFormat.OpenXml.Wordprocessing.StyleValues.Table)) existing.Default = null;
        styles.Append(new DocumentFormat.OpenXml.Wordprocessing.Style(
            new DocumentFormat.OpenXml.Wordprocessing.StyleName { Val = "Bold Grid" },
            new DocumentFormat.OpenXml.Wordprocessing.StyleRunProperties(new DocumentFormat.OpenXml.Wordprocessing.Bold())) {
            Type = DocumentFormat.OpenXml.Wordprocessing.StyleValues.Table, StyleId = "BoldGrid", Default = true
        });
        WordParagraph framed = document.AddParagraph("FRAMED");
        AddFrameParagraph(framed, "around", "text", "margin", x: 1, width: 2880);
        document.AddParagraph("Body text.");

        using PdfPigDocument pdf = RenderFramePdf(document);
        var word = FindFrameWord(pdf, "FRAMED");
        Assert.DoesNotContain("Bold", word.Letters[0].FontName ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SaveAsPdf_FrameWithTextBoxKeepsLaterListMarkers() {
        // A text box inside the first framed paragraph must not shift which source each later
        // framed paragraph resolves its list marker from.
        using WordDocument document = CreateFrameDocument();
        WordParagraph boxed = document.AddParagraph("BOXHOST");
        boxed.AddTextBox("INNERBOX", WordImageTextWrapping.Square);
        AddFrameParagraph(boxed, "around", "text", "margin", x: 1, width: 2880);
        WordList list = document.AddList(WordListStyle.LowerLetterWithDot);
        WordParagraph item = list.AddItem("FRAMEDITEM");
        AddFrameParagraph(item, "around", "text", "margin", x: 1, width: 2880);
        document.AddParagraph("Body text.");

        using PdfPigDocument pdf = RenderFramePdf(document);
        var words = pdf.GetPage(1).GetWords().ToList();
        var framedItem = words.Single(word => word.Text == "FRAMEDITEM");
        Assert.Contains(words, word => word.Text == "a." && Math.Abs(word.Letters[0].StartBaseLine.Y - framedItem.Letters[0].StartBaseLine.Y) < .5D);
    }

    [Fact]
    public void SaveAsPdf_FloatingFrameWithPageFootnotesStillExports() {
        using WordDocument document = CreateFrameDocument();
        AddFrameParagraph(document.AddParagraph("FRAMED"), "around", "text", "margin", x: 1, width: 2880);
        document.AddParagraph("Body with a note.").AddFootNote("NOTETEXT");

        using PdfPigDocument pdf = RenderFramePdf(document);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("FRAMED", text);
        Assert.Contains("NOTETEXT", text);
    }

    [Fact]
    public void SaveAsPdf_FramedHeadingKeepsItsOutlineEntry() {
        using WordDocument document = CreateFrameDocument();
        WordParagraph heading = document.AddParagraph("FRAMEDHEADING");
        heading.Style = WordParagraphStyles.Heading1;
        AddFrameParagraph(heading, "around", "text", "margin", x: 1, width: 2880);
        document.AddParagraph("Body text.");

        using PdfPigDocument pdf = RenderFramePdf(document);
        Assert.True(pdf.TryGetBookmarks(out var bookmarks));
        Assert.Contains(bookmarks.GetNodes(), node => node.Title.Contains("FRAMEDHEADING"));
    }

    [Theory]
    [InlineData("smartart")]
    [InlineData("equation")]
    [InlineData("textbox")]
    [InlineData("floating-picture")]
    [InlineData("footnote")]
    [InlineData("page-field")]
    [InlineData("split-page-field")]
    public void SaveAsPdf_FrameWithContentTheFrameCannotDrawKeepsOrdinaryRendering(string content) {
        // Content outside the frame projection keeps the ordinary paragraph route, which draws
        // it, and the frame approximation is reported.
        using WordDocument document = CreateFrameDocument();
        WordParagraph framed = document.AddParagraph("FRAMED ");
        switch (content) {
            case "smartart":
                framed._paragraph.Append(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 914400, Cy = 914400 },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = 9001U, Name = "Diagram" },
                        new DocumentFormat.OpenXml.Drawing.Graphic(new DocumentFormat.OpenXml.Drawing.GraphicData {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
                        })))));
                break;
            case "equation":
                framed._paragraph.Append(new DocumentFormat.OpenXml.Math.OfficeMath(new DocumentFormat.OpenXml.Math.Run(
                    new DocumentFormat.OpenXml.Math.Text("x"))));
                break;
            case "textbox":
                framed.AddTextBox("INNERBOX", WordImageTextWrapping.Square);
                break;
            case "floating-picture":
                framed.AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), OfficeIMO.OfficeLength.Pixels(32),
                    OfficeIMO.OfficeLength.Pixels(32), new WordImageOptions { Wrapping = WordImageTextWrapping.Square });
                break;
            case "footnote":
                framed.AddFootNote("NOTE");
                break;
            case "page-field":
                framed.AddField(WordFieldType.Page);
                break;
            case "split-page-field":
                // The instruction " PAGE " is split across two runs, as Word often writes it.
                framed._paragraph.Append(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.FieldChar { FieldCharType = DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.Begin }),
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.FieldCode(" PA") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.FieldCode("GE ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.FieldChar { FieldCharType = DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.Separate }),
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("1")),
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.FieldChar { FieldCharType = DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.End }));
                break;
        }
        AddFrameParagraph(framed, "around", "text", "margin", x: 1, width: 2880);
        document.AddParagraph("Body text.");

        var result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.Contains(result.Warnings, warning => warning.Code == "native-paragraph-frame-flattened");
    }

    [Fact]
    public void SaveAsPdf_OutsideHeaderFrameMirrorsOnEvenPages() {
        // Word puts an outside-aligned header frame at the right margin on odd pages and
        // at the left margin on even pages.
        using WordDocument document = CreateFrameDocument();
        AddFrameParagraph(document.PageSetup.Header.AddParagraph("SIDEFRAME"), "around", "text", "margin", y: 1, width: 1440);
        var frame = document.PageSetup.Header._header!.Descendants<DocumentFormat.OpenXml.Wordprocessing.FrameProperties>().Single();
        frame.XAlign = DocumentFormat.OpenXml.Wordprocessing.HorizontalAlignmentValues.Outside;
        document.PageSetup.Header.AddParagraph("ANCHOR");
        document.AddParagraph("Page one.").AddBreak(WordBreakType.Page);
        document.AddParagraph("Page two.");

        using PdfPigDocument pdf = RenderFramePdf(document);
        double OddLeft() => pdf.GetPage(1).GetWords().Single(word => word.Text == "SIDEFRAME").BoundingBox.Left;
        double EvenLeft() => pdf.GetPage(2).GetWords().Single(word => word.Text == "SIDEFRAME").BoundingBox.Left;
        Assert.Equal(FramePageWidth - FrameMargin - 72D, OddLeft(), 0);
        Assert.Equal(FrameMargin, EvenLeft(), 0);
    }

    [Fact]
    public void SaveAsPdf_TallFooterFrameStaysOnThePageWithoutRaisingTheFooter() {
        // Word lets the footer keep its height and moves a frame that would cross the page
        // edge up until it fits.
        using WordDocument document = CreateFrameDocument();
        document.Sections[0].PageSetup.Margins.FooterDistance = OfficeIMO.OfficeLength.Points(18D);
        WordParagraph framed = document.PageSetup.Footer.AddParagraph("TALLFRAME");
        framed.LineSpacingAfter = OfficeIMO.OfficeLength.Points(0D);
        AddFrameParagraph(framed, "around", "text", "margin", x: 1, y: 1, width: 2880);
        var frame = document.PageSetup.Footer._footer!.Descendants<DocumentFormat.OpenXml.Wordprocessing.FrameProperties>().Single();
        frame.Height = 800U;
        frame.HeightType = DocumentFormat.OpenXml.Wordprocessing.HeightRuleValues.Exact;
        document.PageSetup.Footer.AddParagraph("FOOTANCHOR").LineSpacingAfter = OfficeIMO.OfficeLength.Points(0D);
        document.AddParagraph("Body.");
        double alone = MeasureFooterAnchorWithoutFrame();

        using PdfPigDocument pdf = RenderFramePdf(document);
        var top = FindFrameWord(pdf, "TALLFRAME");
        Assert.Equal(alone, FindFrameWord(pdf, "FOOTANCHOR").Letters[0].StartBaseLine.Y, 1);
        // The 40 pt frame ends at the page edge, so its first line sits about 28 pt up.
        Assert.InRange(top.Letters[0].StartBaseLine.Y, 26D, 31D);
        Assert.Equal(FrameMargin + 144D, FindFrameWord(pdf, "FOOTANCHOR").BoundingBox.Left, 0);

        double MeasureFooterAnchorWithoutFrame() {
            using WordDocument plain = CreateFrameDocument();
            plain.Sections[0].PageSetup.Margins.FooterDistance = OfficeIMO.OfficeLength.Points(18D);
            plain.PageSetup.Footer.AddParagraph("FOOTANCHOR").LineSpacingAfter = OfficeIMO.OfficeLength.Points(0D);
            plain.AddParagraph("Body.");
            using PdfPigDocument plainPdf = RenderFramePdf(plain);
            return FindFrameWord(plainPdf, "FOOTANCHOR").Letters[0].StartBaseLine.Y;
        }
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
