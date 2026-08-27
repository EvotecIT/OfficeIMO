using System.IO;
using System.Linq;
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
}
