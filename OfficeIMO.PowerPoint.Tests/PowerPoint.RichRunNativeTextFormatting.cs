using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public class PowerPointRichRunNativeTextFormattingTests {
    [Fact]
    public void RichRunNativeStylesRoundTripWithoutBooleanFlattening() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        try {
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                PowerPointTextRun authoredRun = presentation.AddSlide().AddTextBox("Styled")
                    .Paragraphs.Single().Runs.Single();
                authoredRun.Bold = true;
                authoredRun.Italic = true;
                authoredRun.UnderlineStyle = PowerPointUnderlineStyle.WavyDouble;
                authoredRun.StrikeStyle = PowerPointStrikeStyle.Double;
                authoredRun.Capitalization = PowerPointCapitalization.SmallCaps;
                authoredRun.SetSuperscript();
                authoredRun.FontName = "Aptos";
                authoredRun.FontSizePoints = 14.5D;
                authoredRun.Color = "336699";
                authoredRun.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.ToggleCase);
                presentation.Save();
            }

            using PowerPointPresentation reopened = PowerPointPresentation.Load(path);
            PowerPointTextRun actual = reopened.Slides.Single().TextBoxes.Single()
                .Paragraphs.Single().Runs.Single();
            Assert.Equal("sTYLED", actual.Text);
            Assert.True(actual.Bold);
            Assert.True(actual.Italic);
            Assert.True(actual.Underline);
            Assert.Equal(PowerPointUnderlineStyle.WavyDouble, actual.UnderlineStyle);
            Assert.True(actual.Strikethrough);
            Assert.Equal(PowerPointStrikeStyle.Double, actual.StrikeStyle);
            Assert.Equal(PowerPointCapitalization.SmallCaps, actual.Capitalization);
            Assert.Equal(30D, actual.BaselinePercent);
            Assert.Equal("Aptos", actual.FontName);
            Assert.Equal(14.5D, actual.FontSizePoints);
            Assert.Equal("336699", actual.Color);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void BaselineRejectsOutOfRangeValues() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointTextRun run = presentation.AddSlide().AddTextBox("Styled")
            .Paragraphs.Single().Runs.Single();

        Assert.Throws<ArgumentOutOfRangeException>(() => run.BaselinePercent = 100.1D);
        Assert.Throws<ArgumentOutOfRangeException>(() => run.BaselinePercent = double.NaN);
    }
}
