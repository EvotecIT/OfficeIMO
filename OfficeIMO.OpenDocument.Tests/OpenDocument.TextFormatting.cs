using System.IO;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentTextFormattingTests {
    [Fact]
    public void OdtSpanNativeDecorationScriptAndCaseStylesRoundTrip() {
        OdtDocument document = OdtDocument.Create();
        OdtSpan span = document.AddParagraph().AddSpan("Styled");
        span.Bold = true;
        span.Italic = true;
        span.UnderlineStyle = OdfTextDecorationStyle.Wave;
        span.UnderlineType = OdfTextDecorationType.Double;
        span.LineThroughStyle = OdfTextDecorationStyle.Dotted;
        span.LineThroughType = OdfTextDecorationType.Single;
        span.TextPosition = OdfTextPosition.Superscript;
        span.TextTransform = OdfTextTransform.Uppercase;
        span.SmallCaps = true;
        span.FontFamily = "Liberation Sans";
        span.FontSize = OdfLength.Parse("14pt");
        span.Color = OdfColor.Parse("#336699");
        span.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.ToggleCase);

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        OdtSpan actual = reopened.Paragraphs.Single().Spans.Single();
        Assert.Equal("sTYLED", actual.Text);
        Assert.True(actual.Bold);
        Assert.True(actual.Italic);
        Assert.True(actual.Underline);
        Assert.Equal(OdfTextDecorationStyle.Wave, actual.UnderlineStyle);
        Assert.Equal(OdfTextDecorationType.Double, actual.UnderlineType);
        Assert.True(actual.StrikeThrough);
        Assert.Equal(OdfTextDecorationStyle.Dotted, actual.LineThroughStyle);
        Assert.Equal(OdfTextDecorationType.Single, actual.LineThroughType);
        Assert.Equal(OdfTextPosition.Superscript, actual.TextPosition);
        Assert.Equal(OdfTextTransform.Uppercase, actual.TextTransform);
        Assert.True(actual.SmallCaps);
        Assert.Equal("Liberation Sans", actual.FontFamily);
        Assert.Equal(OdfLength.Parse("14pt"), actual.FontSize);
        Assert.Equal(OdfColor.Parse("#336699"), actual.Color);
    }

    [Fact]
    public void OdpRunNativeDecorationScriptAndCaseStylesRoundTrip() {
        OdpPresentation document = OdpPresentation.Create();
        OdpTextBox textBox = document.AddSlide("Text")
            .AddTextBox(OdfRect.FromCentimeters(1, 1, 10, 3), null, "Text");
        OdpRun run = textBox.AddParagraph().AddRun("Styled");
        run.UnderlineStyle = OdfTextDecorationStyle.DotDash;
        run.UnderlineType = OdfTextDecorationType.Double;
        run.LineThroughStyle = OdfTextDecorationStyle.Wave;
        run.LineThroughType = OdfTextDecorationType.Single;
        run.TextPosition = OdfTextPosition.Subscript;
        run.TextTransform = OdfTextTransform.Lowercase;
        run.SmallCaps = true;

        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(document.ToBytes()));
        OdpRun actual = Assert.Single(Assert.IsType<OdpTextBox>(Assert.Single(reopened.Slides[0].Shapes))
            .Paragraphs.Single().Runs);
        Assert.Equal(OdfTextDecorationStyle.DotDash, actual.UnderlineStyle);
        Assert.Equal(OdfTextDecorationType.Double, actual.UnderlineType);
        Assert.Equal(OdfTextDecorationStyle.Wave, actual.LineThroughStyle);
        Assert.Equal(OdfTextDecorationType.Single, actual.LineThroughType);
        Assert.Equal(OdfTextPosition.Subscript, actual.TextPosition);
        Assert.Equal(OdfTextTransform.Lowercase, actual.TextTransform);
        Assert.True(actual.SmallCaps);
    }

    [Fact]
    public void OdsCellNativeDecorationScriptAndCaseStylesRoundTrip() {
        OdsDocument document = OdsDocument.Create();
        OdsCell cell = document.AddSheet("Text").Cell(0, 0);
        cell.SetString("Styled");
        cell.UnderlineStyle = OdfTextDecorationStyle.Dotted;
        cell.UnderlineType = OdfTextDecorationType.Double;
        cell.LineThroughStyle = OdfTextDecorationStyle.Dash;
        cell.LineThroughType = OdfTextDecorationType.Single;
        cell.TextPosition = OdfTextPosition.Superscript;
        cell.TextTransform = OdfTextTransform.Capitalize;
        cell.SmallCaps = true;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(document.ToBytes()));
        OdsCell actual = reopened.Sheets.Single().Cell(0, 0);
        Assert.Equal(OdfTextDecorationStyle.Dotted, actual.UnderlineStyle);
        Assert.Equal(OdfTextDecorationType.Double, actual.UnderlineType);
        Assert.Equal(OdfTextDecorationStyle.Dash, actual.LineThroughStyle);
        Assert.Equal(OdfTextDecorationType.Single, actual.LineThroughType);
        Assert.Equal(OdfTextPosition.Superscript, actual.TextPosition);
        Assert.Equal(OdfTextTransform.Capitalize, actual.TextTransform);
        Assert.True(actual.SmallCaps);
    }

    [Fact]
    public void ParagraphCaseTransformsPreserveRunsAndHyperlinks() {
        OdtDocument odt = OdtDocument.Create();
        OdtParagraph odtParagraph = odt.AddParagraph().AddText("Plain ");
        OdtSpan odtSpan = odtParagraph.AddSpan("Styled ");
        odtSpan.Bold = true;
        odtParagraph.AddHyperlink("Link", "https://example.test/odt");
        odtParagraph.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.Uppercase, CultureInfo.InvariantCulture);

        OdtDocument reopenedOdt = OdtDocument.Load(new MemoryStream(odt.ToBytes()));
        OdtParagraph actualOdt = Assert.Single(reopenedOdt.Paragraphs);
        Assert.Equal("PLAIN STYLED LINK", actualOdt.Text);
        Assert.True(Assert.Single(actualOdt.Spans).Bold);
        Assert.Equal("https://example.test/odt", Assert.Single(actualOdt.Hyperlinks).Href);

        OdpPresentation odp = OdpPresentation.Create();
        OdpParagraph odpParagraph = odp.AddSlide("Case")
            .AddTextBox(OdfRect.FromCentimeters(1, 1, 10, 3), null, "Case")
            .AddParagraph("Plain ");
        OdpRun odpRun = odpParagraph.AddRun("Styled ");
        odpRun.Italic = true;
        odpParagraph.AddHyperlink("Link", "https://example.test/odp");
        odpParagraph.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.Uppercase, CultureInfo.InvariantCulture);

        OdpPresentation reopenedOdp = OdpPresentation.Load(new MemoryStream(odp.ToBytes()));
        OdpParagraph actualOdp = Assert.IsType<OdpTextBox>(Assert.Single(reopenedOdp.Slides[0].Shapes))
            .Paragraphs.Single();
        Assert.Equal("PLAIN STYLED LINK", actualOdp.Text);
        Assert.True(Assert.Single(actualOdp.Runs).Italic);
        Assert.Equal("https://example.test/odp", Assert.Single(actualOdp.InlineNodes,
            node => node.Kind == OdpInlineNodeKind.Hyperlink).Hyperlink!.Href);
    }

    [Fact]
    public void ContextualCaseTransformsContinueAcrossInlineNodeBoundaries() {
        OdtDocument odt = OdtDocument.Create();
        OdtParagraph odtParagraph = odt.AddParagraph().AddText("mi");
        OdtSpan odtSpan = odtParagraph.AddSpan("XED ti");
        odtSpan.Bold = true;
        odtParagraph.AddHyperlink("TLE", "https://example.test/title");
        odtParagraph.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.TitleCase, CultureInfo.InvariantCulture);

        OdtParagraph actualOdt = Assert.Single(OdtDocument.Load(new MemoryStream(odt.ToBytes())).Paragraphs);
        Assert.Equal("Mixed Title", actualOdt.Text);
        Assert.True(Assert.Single(actualOdt.Spans).Bold);
        Assert.Equal("https://example.test/title", Assert.Single(actualOdt.Hyperlinks).Href);

        OdpPresentation odp = OdpPresentation.Create();
        OdpParagraph odpParagraph = odp.AddSlide("Case")
            .AddTextBox(OdfRect.FromCentimeters(1, 1, 10, 3), null, "Case")
            .AddParagraph("he");
        OdpRun odpRun = odpParagraph.AddRun("LLO. w");
        odpRun.Italic = true;
        odpParagraph.AddHyperlink("ORLD", "https://example.test/sentence");
        odpParagraph.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.SentenceCase, CultureInfo.InvariantCulture);

        OdpParagraph actualOdp = Assert.IsType<OdpTextBox>(Assert.Single(
            OdpPresentation.Load(new MemoryStream(odp.ToBytes())).Slides[0].Shapes)).Paragraphs.Single();
        Assert.Equal("Hello. World", actualOdp.Text);
        Assert.True(Assert.Single(actualOdp.Runs).Italic);
        Assert.Equal("https://example.test/sentence", Assert.Single(actualOdp.InlineNodes,
            node => node.Kind == OdpInlineNodeKind.Hyperlink).Hyperlink!.Href);
    }

    [Fact]
    public void ParagraphCaseTransformsExcludeAnnotationsAndEmbeddedObjectMetadata() {
        OdtDocument document = OdtDocument.Create();
        OdtParagraph paragraph = document.AddParagraph().AddText("hELLO ");
        var annotation = new XElement(OdfNamespaces.Office + "annotation",
            new XElement(OdfNamespaces.Dc + "creator", "aUTHOR"),
            new XElement(OdfNamespaces.Text + "p", "nOTE."));
        var embeddedObject = new XElement(OdfNamespaces.Draw + "object",
            new XElement(OdfNamespaces.Office + "binary-data", "mETADATA."));
        paragraph.Element.Add(annotation, embeddedObject);
        OdtSpan trailing = paragraph.AddSpan("wORLD");

        paragraph.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.SentenceCase, CultureInfo.InvariantCulture);

        Assert.Equal("Hello", Assert.IsType<XText>(paragraph.Element.FirstNode).Value);
        Assert.Equal(OdfNamespaces.Text + "s", paragraph.Element.Elements().First().Name);
        Assert.Equal("world", trailing.Text);
        Assert.Equal("aUTHOR", annotation.Element(OdfNamespaces.Dc + "creator")!.Value);
        Assert.Equal("nOTE.", annotation.Element(OdfNamespaces.Text + "p")!.Value);
        Assert.Equal("mETADATA.", embeddedObject.Descendants(OdfNamespaces.Office + "binary-data").Single().Value);
    }

    [Fact]
    public void PercentageTextPositionsMapToNativeBaselineValues() {
        OdtDocument document = OdtDocument.Create();
        OdtParagraph paragraph = document.AddParagraph();
        OdtSpan raised = paragraph.AddSpan("raised");
        OdtSpan lowered = paragraph.AddSpan("lowered");
        OdtSpan normal = paragraph.AddSpan("normal");
        raised.TextPosition = OdfTextPosition.Normal;
        lowered.TextPosition = OdfTextPosition.Normal;
        normal.TextPosition = OdfTextPosition.Superscript;

        OdfStyle raisedStyle = document.Styles.FindInPart(OdfStyleFamily.Text, raised.StyleName!, "content.xml")!;
        OdfStyle loweredStyle = document.Styles.FindInPart(OdfStyleFamily.Text, lowered.StyleName!, "content.xml")!;
        OdfStyle normalStyle = document.Styles.FindInPart(OdfStyleFamily.Text, normal.StyleName!, "content.xml")!;
        raisedStyle.TextProperties!.SetAttributeValue(OdfNamespaces.Style + "text-position", "33% 58%");
        loweredStyle.TextProperties!.SetAttributeValue(OdfNamespaces.Style + "text-position", "-25% 58%");
        normalStyle.TextProperties!.SetAttributeValue(OdfNamespaces.Style + "text-position", "0% 100%");

        OdtParagraph actual = Assert.Single(OdtDocument.Load(new MemoryStream(document.ToBytes())).Paragraphs);
        Assert.Equal(OdfTextPosition.Superscript, actual.Spans[0].TextPosition);
        Assert.Equal(OdfTextPosition.Subscript, actual.Spans[1].TextPosition);
        Assert.Equal(OdfTextPosition.Normal, actual.Spans[2].TextPosition);
    }
}
