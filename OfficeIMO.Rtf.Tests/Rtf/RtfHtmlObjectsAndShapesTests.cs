using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlObjectsAndShapesTests {
    [Fact]
    public void RtfDocument_ToHtml_Renders_Inline_Object_Metadata_And_RoundTrips() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Before ");
        RtfObject rtfObject = paragraph.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3, 255 });
        rtfObject.ClassName = "Package";
        rtfObject.Name = "Attachment";
        rtfObject.Width = 100;
        rtfObject.Height = 200;
        rtfObject.ScaleX = 75;
        rtfObject.ScaleY = 80;
        rtfObject.Result.AddText("Display").SetBold();
        paragraph.AddText(" after");

        string html = document.ToHtml(RtfToHtmlOptions.CreateRoundTripProfile());

        Assert.Contains("data-officeimo-rtf-object=\"embedded\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-object-class=\"Package\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-object-name=\"Attachment\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-object-data=\"AQID/w==\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-object-width=\"100\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-object-result=\"", html, StringComparison.Ordinal);

        RtfParagraph roundTripParagraph = Assert.Single(HtmlConversionDocument.Parse(html).ToRtfDocument().Paragraphs);
        Assert.Equal("Before Display after", roundTripParagraph.ToPlainText());
        RtfObject roundTripObject = Assert.IsType<RtfObject>(roundTripParagraph.Inlines[1]);
        Assert.Equal(RtfObjectKind.Embedded, roundTripObject.Kind);
        Assert.Equal("Package", roundTripObject.ClassName);
        Assert.Equal("Attachment", roundTripObject.Name);
        Assert.Equal(new byte[] { 1, 2, 3, 255 }, roundTripObject.Data);
        Assert.Equal(100, roundTripObject.Width);
        Assert.Equal(200, roundTripObject.Height);
        Assert.Equal(75, roundTripObject.ScaleX);
        Assert.Equal(80, roundTripObject.ScaleY);
        Assert.Contains(roundTripObject.Result.Runs, run => run.Text == "Display" && run.Bold);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Block_Object_Metadata() {
        const string html = "<div data-officeimo-rtf-object=\"linked\" data-officeimo-rtf-object-class=\"Package\" data-officeimo-rtf-object-name=\"Attachment\" data-officeimo-rtf-object-data=\"AQID\" data-officeimo-rtf-object-width=\"100\" data-officeimo-rtf-object-height=\"200\"></div>";

        RtfDocument document = HtmlConversionDocument.Parse(html).ToRtfDocument();

        RtfObject rtfObject = Assert.IsType<RtfObject>(Assert.Single(document.Blocks));
        Assert.Equal(RtfObjectKind.Linked, rtfObject.Kind);
        Assert.Equal("Package", rtfObject.ClassName);
        Assert.Equal("Attachment", rtfObject.Name);
        Assert.Equal(new byte[] { 1, 2, 3 }, rtfObject.Data);
        Assert.Equal(100, rtfObject.Width);
        Assert.Equal(200, rtfObject.Height);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\object\objlink\objw100\objh200{\*\objclass Package}{\*\objname Attachment}{\*\objdata 010203}}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Shape_Metadata_And_RoundTrips() {
        RtfDocument document = RtfDocument.Create();
        RtfShape shape = document.AddShape();
        shape.AddInstruction("shpleft", 100);
        shape.AddInstruction("shptop", 200);
        shape.AddInstruction("shpright", 2100);
        shape.AddInstruction("shpbottom", 900);
        shape.AddProperty("shapeType", "202");
        shape.AddProperty("fLine", "0");
        RtfParagraph textBox = shape.AddTextBoxParagraph("Text ");
        textBox.AddText("box").SetItalic();

        string html = document.ToHtml(RtfToHtmlOptions.CreateRoundTripProfile());

        Assert.Contains("data-officeimo-rtf-shape=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-shape-instructions=\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-shape-properties=\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-shape-text=\"", html, StringComparison.Ordinal);

        RtfShape roundTripShape = Assert.IsType<RtfShape>(Assert.Single(HtmlConversionDocument.Parse(html).ToRtfDocument().Blocks));
        Assert.Contains(roundTripShape.Instructions, instruction => instruction.Name == "shpleft" && instruction.Parameter == 100);
        Assert.Contains(roundTripShape.Instructions, instruction => instruction.Name == "shpbottom" && instruction.Parameter == 900);
        Assert.Contains(roundTripShape.Properties, property => property.Name == "shapeType" && property.Value == "202");
        Assert.Contains(roundTripShape.Properties, property => property.Name == "fLine" && property.Value == "0");
        RtfParagraph roundTripText = Assert.Single(roundTripShape.TextBoxParagraphs);
        Assert.Equal("Text box", roundTripText.ToPlainText());
        Assert.Contains(roundTripText.Runs, run => run.Text == "box" && run.Italic);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Inline_Shape_Text_Without_Block_Paragraphs() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Before ");
        RtfShape shape = paragraph.AddShape();
        shape.AddTextBoxParagraph("One");
        shape.AddTextBoxParagraph("Two");
        paragraph.AddText(" after");

        string html = document.ToHtml();

        Assert.Contains("<span class=\"rtf-shape-text\">One<br>Two</span>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<span class=\"rtf-shape-text\"><p", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_ToRtfDocument_Applies_Active_Url_Policy_To_Encoded_Object_And_Shape_Content() {
        string encodedObjectResult = Convert.ToBase64String(Encoding.UTF8.GetBytes("<p><a href=\"file:///C:/secret-object\">Object link</a></p>"));
        string encodedShapeText = Convert.ToBase64String(Encoding.UTF8.GetBytes("<p><a href=\"file:///C:/secret-shape\">Shape link</a></p>"));
        string html = "<div data-officeimo-rtf-object=\"embedded\" data-officeimo-rtf-object-result=\"" + encodedObjectResult + "\"></div>" +
            "<div data-officeimo-rtf-shape=\"true\" data-officeimo-rtf-shape-text=\"" + encodedShapeText + "\"></div>";
        var options = HtmlToRtfOptions.CreateUntrustedHtmlProfile();
        options.UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult(options);
        RtfDocument document = result.Value;

        RtfObject rtfObject = Assert.IsType<RtfObject>(document.Blocks[0]);
        Assert.All(rtfObject.Result.Runs, run => Assert.Null(run.Hyperlink));
        RtfShape shape = Assert.IsType<RtfShape>(document.Blocks[1]);
        Assert.All(Assert.Single(shape.TextBoxParagraphs).Runs, run => Assert.Null(run.Hyperlink));
        Assert.Equal("Object link", rtfObject.Result.ToPlainText());
        Assert.Equal("Shape link", Assert.Single(shape.TextBoxParagraphs).ToPlainText());
        Assert.Equal(2, result.RtfDiagnostics.Count(diagnostic => diagnostic.Code == "HtmlRtfHyperlinkRejected"));
        Assert.Throws<RtfConversionLossException>(() => result.RtfReport.RequireNoLoss());
    }

    [Fact]
    public void Html_ToRtfDocument_SourcePolicyDiagnosticsDoNotReintroduceExecutableContent() {
        const string html = "<p onclick='alert(1)'>Visible<script>Hidden script</script></p>";

        RtfDocument document = HtmlConversionDocument.Parse(html).ToRtfDocument();

        Assert.Equal("Visible", Assert.Single(document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Html_ToRtfDocument_Applies_Active_Node_Limit_To_Encoded_Object_Content() {
        string nestedHtml = "<p>" + string.Concat(Enumerable.Repeat("<span>x</span>", 16)) + "</p>";
        string encodedObjectResult = Convert.ToBase64String(Encoding.UTF8.GetBytes(nestedHtml));
        string html = "<div data-officeimo-rtf-object=\"embedded\" data-officeimo-rtf-object-result=\"" + encodedObjectResult + "\"></div>";
        var options = new HtmlToRtfOptions { MaxHtmlNodes = 10 };

        HtmlRtfConversionLimitException exception = Assert.Throws<HtmlRtfConversionLimitException>(() => HtmlConversionDocument.Parse(html).ToRtfDocument(options));

        Assert.Equal("MaxHtmlNodes", exception.LimitSource);
    }
}
