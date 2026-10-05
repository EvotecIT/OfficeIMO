using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public class PowerPointRunFormattingActionOrderTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void FormattingAndActionsRemainValidInEitherOrder(bool tableCell, bool actionsFirst) {
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

        if (actionsFirst) run.SetHyperlink("https://example.test/linked-label");
        run.SetMouseOverStopsSound(true);
        run.FontName = "Aptos";
        run.HighlightColor = "FFFF00";
        run.Color = "FFFFFF";
        // Replacing the link must preserve the formatting and mouse-over action as well.
        run.SetHyperlink("https://example.test/final-label");

        using var stream = new MemoryStream();
        presentation.Save(stream);
        stream.Position = 0;
        using var reopened = PowerPointPresentation.Load(stream);
        Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
        PowerPointSlide actualSlide = reopened.Slides.Single();
        PowerPointTextRun actual = tableCell
            ? actualSlide.Tables.Single().GetCell(0, 0).Runs.Single()
            : actualSlide.TextBoxes.Single().Paragraphs.Single().Runs.Single();
        Assert.Equal("Linked label", actual.Text);
        Assert.Equal("Aptos", actual.FontName);
        Assert.Equal("FFFF00", actual.HighlightColor);
        Assert.Equal("FFFFFF", actual.Color);
        Assert.Equal("https://example.test/final-label", actual.Hyperlink?.AbsoluteUri);
        Assert.True(actual.HasMouseOverInteraction);
        Assert.True(actual.MouseOverStopsSound);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CellAndTextBoxFormattingPreservesLinkedRuns(bool tableCell) {
        using var presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTableCell? cell = null;
        PowerPointTextBox? textBox = null;
        PowerPointTextRun run;
        if (tableCell) {
            cell = slide.AddTablePoints(1, 1, 20, 20, 300, 80).GetCell(0, 0);
            cell.Text = "Linked label";
            run = cell.Runs.Single();
        } else {
            textBox = slide.AddTextBox("Linked label");
            run = textBox.Paragraphs.Single().Runs.Single();
        }
        run.SetHyperlink("https://example.test/linked-label");
        run.SetMouseOverStopsSound(true);
        if (cell != null) {
            cell.FontName = "Aptos";
            cell.Color = "FFFFFF";
        } else {
            textBox!.FontName = "Aptos";
            textBox.Color = "FFFFFF";
        }

        using var stream = new MemoryStream();
        presentation.Save(stream);
        stream.Position = 0;
        using var reopened = PowerPointPresentation.Load(stream);
        Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
        PowerPointSlide actualSlide = reopened.Slides.Single();
        PowerPointTextRun actual = tableCell
            ? actualSlide.Tables.Single().GetCell(0, 0).Runs.Single()
            : actualSlide.TextBoxes.Single().Paragraphs.Single().Runs.Single();
        Assert.Equal("Linked label", actual.Text);
        Assert.Equal("Aptos", actual.FontName);
        Assert.Equal("FFFFFF", actual.Color);
        Assert.Equal("https://example.test/linked-label", actual.Hyperlink?.AbsoluteUri);
        Assert.True(actual.MouseOverStopsSound);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitColorReplacesImportedFillWithoutChangingScriptFonts(bool gradient) {
        using var presentation = PowerPointPresentation.Create();
        PowerPointTextRun run = presentation.AddSlide().AddTextBox("Imported label")
            .Paragraphs.Single().Runs.Single();
        run.SetHyperlink("https://example.test/imported-label");
        A.RunProperties properties = presentation.OpenXmlDocument.PresentationPart!.SlideParts
            .Single().Slide.Descendants<A.RunProperties>().Single();
        properties.AddChild(new A.LatinFont { Typeface = "Aptos" }, true);
        properties.AddChild(new A.EastAsianFont { Typeface = "Noto Sans CJK JP" }, true);
        properties.AddChild(new A.ComplexScriptFont { Typeface = "Noto Sans Arabic" }, true);
        if (gradient) {
            properties.AddChild(new A.GradientFill(new A.GradientStopList(
                new A.GradientStop(new A.RgbColorModelHex { Val = "123456" }) { Position = 0 },
                new A.GradientStop(new A.RgbColorModelHex { Val = "654321" }) { Position = 100000 })), true);
        } else {
            properties.AddChild(new A.NoFill(), true);
        }
        Assert.Empty(new OpenXmlValidator().Validate(presentation.OpenXmlDocument));
        using var source = new MemoryStream();
        presentation.Save(source);
        source.Position = 0;
        using var imported = PowerPointPresentation.Load(source);
        imported.Slides.Single().TextBoxes.Single().Paragraphs.Single().Runs.Single().Color = "FFFFFF";
        using var output = new MemoryStream();
        imported.Save(output);
        output.Position = 0;
        using var reopened = PowerPointPresentation.Load(output);
        Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
        A.RunProperties actual = reopened.OpenXmlDocument.PresentationPart!.SlideParts
            .Single().Slide.Descendants<A.RunProperties>().Single();
        Assert.Equal("FFFFFF", actual.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
        Assert.Null(actual.GetFirstChild<A.GradientFill>());
        Assert.Null(actual.GetFirstChild<A.NoFill>());
        Assert.Equal("Aptos", actual.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
        Assert.Equal("Noto Sans CJK JP", actual.GetFirstChild<A.EastAsianFont>()?.Typeface?.Value);
        Assert.Equal("Noto Sans Arabic", actual.GetFirstChild<A.ComplexScriptFont>()?.Typeface?.Value);
        Assert.Equal("https://example.test/imported-label", reopened.Slides.Single().TextBoxes.Single()
            .Paragraphs.Single().Runs.Single().Hyperlink?.AbsoluteUri);
    }
}
