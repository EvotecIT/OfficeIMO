using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfTextCaseTests {
    [Fact]
    public void TransformTextCasePreservesRunFormatting() {
        RtfDocument document = RtfDocument.Create();
        RtfRun run = document.AddParagraph().AddText("Styled");
        run.Bold = true;
        run.Italic = true;
        run.UnderlineStyle = RtfUnderlineStyle.DoubleWave;
        run.DoubleStrike = true;
        run.VerticalPosition = RtfVerticalPosition.Superscript;
        run.TransformTextCase(OfficeTextCase.ToggleCase);

        RtfRun actual = document.Paragraphs.Single().Runs.Single();
        Assert.Equal("sTYLED", actual.Text);
        Assert.True(actual.Bold);
        Assert.True(actual.Italic);
        Assert.Equal(RtfUnderlineStyle.DoubleWave, actual.UnderlineStyle);
        Assert.True(actual.DoubleStrike);
        Assert.Equal(RtfVerticalPosition.Superscript, actual.VerticalPosition);
    }
}
