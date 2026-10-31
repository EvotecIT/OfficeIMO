using System;
using System.IO;
using System.Linq;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPowerPointComputedFontsTests {
    [Theory]
    [InlineData("font-family:inherit", "Parent Face")]
    [InlineData("font-family:unset", "Parent Face")]
    [InlineData("font-family:var(--missing-family, inherit)", "Parent Face")]
    [InlineData("", "Parent Face")]
    [InlineData("font-family:var(--family)", "Variable Face")]
    [InlineData("font-family:var(--missing-family, 'Fallback Face')", "Fallback Face")]
    [InlineData("font-family:'inherit'", "inherit")]
    public void OrdinarySpanFontUsesComputedCssInSavedPresentation(string declaration, string expectedFamily) {
        string html = "<body style=\"font-family:'Parent Face';--family:'Variable Face'\">"
            + "<section><h1>Computed font controls</h1><p><span style=\"" + declaration
            + "\">FAMILYMARKER</span></p></section></body>";
        using PowerPointDocument imported = ImportGeneric(html);

        AssertBeforeAndAfterSaving(imported, presentation =>
            AssertFont(presentation, "FAMILYMARKER", expectedFamily));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HeadingAndBodyKeepTheirOwnTextAndFontsInSavedPresentation(bool explicitSection) {
        const string content = "<h1 style=\"font-family:'Heading Face'\"><span>HEADINGMARKER</span></h1>"
            + "<p><span>BODYMARKER</span></p>";
        string html = "<body style=\"font-family:'Body Face'\">"
            + (explicitSection ? "<section>" + content + "</section>" : content) + "</body>";
        using PowerPointDocument imported = ImportGeneric(html);

        AssertBeforeAndAfterSaving(imported, presentation => {
            PowerPointSlide slide = Assert.Single(presentation.Slides);
            Assert.Collection(slide.TextBoxes,
                title => Assert.Equal("HEADINGMARKER", title.Text),
                body => Assert.Equal("BODYMARKER", body.Text));
            AssertFont(presentation, "HEADINGMARKER", "Heading Face");
            AssertFont(presentation, "BODYMARKER", "Body Face");
        });
    }

    [Fact]
    public void DocumentTitleDoesNotBorrowTheFirstBodyRunInSavedPresentation() {
        const string html = "<html><head><title>Deck title</title></head>"
            + "<body style=\"font-family:'Body Face'\"><p><span>BODYMARKER</span></p></body></html>";
        using PowerPointDocument imported = ImportGeneric(html);

        AssertBeforeAndAfterSaving(imported, presentation => {
            PowerPointSlide slide = Assert.Single(presentation.Slides);
            Assert.Collection(slide.TextBoxes,
                title => Assert.Equal("Deck title", title.Text),
                body => Assert.Equal("BODYMARKER", body.Text));
            AssertFont(presentation, "BODYMARKER", "Body Face");
        });
    }

    private static PowerPointDocument ImportGeneric(string html) => HtmlConversionDocument.Parse(html)
        .ToPowerPointDocumentResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic })
        .RequireValue();

    private static void AssertFont(PowerPointDocument presentation, string text, string expectedFamily) {
        PowerPointTextRun run = Assert.Single(presentation.Slides.SelectMany(slide => slide.TextBoxes)
            .SelectMany(box => box.Paragraphs).SelectMany(paragraph => paragraph.Runs),
            candidate => candidate.Text == text);
        Assert.Equal(expectedFamily, run.FontName);
    }

    private static void AssertBeforeAndAfterSaving(
        PowerPointDocument imported, Action<PowerPointDocument> assert) {
        assert(imported);
        using var artifact = new MemoryStream();
        imported.Save(artifact);
        using var input = new MemoryStream(artifact.ToArray());
        using PowerPointDocument reopened = PowerPointDocument.Load(input,
            new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        assert(reopened);
    }
}
