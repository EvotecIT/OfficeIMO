using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using H = DocumentFormat.OpenXml.Office2019.Drawing.HyperLinkColor;

namespace OfficeIMO.Tests;

public class PowerPointHyperlinkTextColorTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ExplicitTextColorRemainsEffectiveAfterLinkCreationAndReplacement(bool tableCell, bool linkFirst) {
        using var presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextRun run;
        if (tableCell) {
            PowerPointTableCell cell = slide.AddTablePoints(1, 1, 20, 20, 300, 80).GetCell(0, 0);
            cell.Text = "Linked label";
            run = cell.Runs.Single();
        } else {
            run = slide.AddTextBox("Linked label").Paragraphs.Single().Runs.Single();
        }
        if (linkFirst) run.SetHyperlink("https://example.test/first");
        run.Color = "FFFFFF";
        run.SetHyperlink("https://example.test/final");
        using var output = new MemoryStream();
        presentation.Save(output);
        output.Position = 0;
        using var reopened = PowerPointPresentation.Load(output);
        Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
        A.HyperlinkOnClick hyperlink = reopened.OpenXmlDocument.PresentationPart!.SlideParts
            .Single().Slide.Descendants<A.HyperlinkOnClick>().Single();
        H.HyperlinkColor color = Assert.Single(hyperlink.Descendants<H.HyperlinkColor>());
        Assert.Equal(H.HyperlinkColorEnum.Tx, color.Val!.Value);
        PowerPointTextRun actual = tableCell
            ? reopened.Slides.Single().Tables.Single().GetCell(0, 0).Runs.Single()
            : reopened.Slides.Single().TextBoxes.Single().Paragraphs.Single().Runs.Single();
        Assert.Equal("FFFFFF", actual.Color);
        Assert.Equal("https://example.test/final", actual.Hyperlink?.AbsoluteUri);
        Assert.Equal("Linked label", actual.Text);
    }

    [Fact]
    public void UncoloredLinksKeepThemeHyperlinkColor() {
        using var presentation = PowerPointPresentation.Create();
        var run = presentation.AddSlide().AddTextBox("Ordinary link").Paragraphs.Single().Runs.Single();
        run.SetHyperlink("https://example.test/source");
        Assert.Empty(presentation.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide
            .Descendants<H.HyperlinkColor>());
    }
    [Fact]
    public void InheritedTextColorPolicyAndThemeOverrideSurviveTargetReplacement() {
        using var presentation = PowerPointPresentation.Create();
        var run = presentation.AddSlide().AddTextBox("Styled link").Paragraphs.Single().Runs.Single();
        run.SetHyperlink("https://example.test/source");
        run.HyperlinkUsesTextColor = true;
        Assert.Null(run.Color);
        run.SetHyperlink("https://example.test/replaced");
        Assert.True(run.HyperlinkUsesTextColor);
        run.Color = "FFFFFF";
        run.Color = null;
        Assert.Null(run.Color);
        Assert.True(run.HyperlinkUsesTextColor);
        run.Color = "FFFFFF";
        run.HyperlinkUsesTextColor = false;
        run.SetHyperlink("https://example.test/theme");
        Assert.False(run.HyperlinkUsesTextColor);
        using var stream = new MemoryStream();
        presentation.Save(stream);
        stream.Position = 0;
        using var reopened = PowerPointPresentation.Load(stream);
        Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
        var actual = reopened.Slides.Single().TextBoxes.Single().Paragraphs.Single().Runs.Single();
        Assert.False(actual.HyperlinkUsesTextColor);
        Assert.Equal("FFFFFF", actual.Color);
        Assert.Equal("https://example.test/theme", actual.Hyperlink!.AbsoluteUri);
    }

    [Fact]
    public void NativeColorExtensionRetainsOtherExtensionsAndSoundSchemaOrder() {
        using var presentation = PowerPointPresentation.Create();
        var run = presentation.AddSlide().AddTextBox("Sound link").Paragraphs.Single().Runs.Single();
        run.SetHyperlink("https://example.test/source");
        A.HyperlinkOnClick hyperlink = presentation.OpenXmlDocument.PresentationPart!.SlideParts
            .Single().Slide.Descendants<A.HyperlinkOnClick>().Single();
        hyperlink.HyperlinkExtensionList = new A.HyperlinkExtensionList(
            new A.HyperlinkExtension { Uri = "{F60F15CC-1C01-4324-90E6-D474220C5BB0}" });
        run.Color = "FFFFFF";
        run.Color = "FFDD00";
        byte[] sound = {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F', 40, 0, 0, 0,
            (byte)'W', (byte)'A', (byte)'V', (byte)'E', (byte)'f', (byte)'m', (byte)'t', (byte)' ',
            16, 0, 0, 0, 1, 0, 1, 0, 0x40, 0x1F, 0, 0, 0x40, 0x1F, 0, 0,
            1, 0, 8, 0, (byte)'d', (byte)'a', (byte)'t', (byte)'a', 4, 0, 0, 0, 0x80, 0x90, 0x70, 0x80
        };
        using (var audio = new MemoryStream(sound)) run.SetClickSound(audio, "Chime");
        run.SetClickStopsSound(true);
        Assert.True(run.HyperlinkUsesTextColor);
        Assert.Single(hyperlink.Descendants<H.HyperlinkColor>());
        Assert.Contains(hyperlink.HyperlinkExtensionList.Elements<A.HyperlinkExtension>(),
            extension => extension.Uri!.Value == "{F60F15CC-1C01-4324-90E6-D474220C5BB0}");
        Assert.True(run.ClickStopsSound);
        Assert.Equal(sound, run.GetClickSoundBytes());
        Assert.Empty(new OpenXmlValidator().Validate(presentation.OpenXmlDocument));
    }

}
