using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote.Tests;

public sealed class NativeMathOperatorSpacingLossTests {
    [Theory]
    [InlineData("block-expression")]
    [InlineData("block-mathml")]
    [InlineData("inline")]
    [InlineData("edited-preserved-inline")]
    public void NativeWriterRejectsOperatorMetadataBeforeWritingDestination(string entryPoint) {
        const string mathMl = "<math><mrow><mi>x</mi><mo lspace='.5em'>+</mo><mi>y</mi></mrow></math>";
        OfficeMathExpression expression = OfficeMathMarkup.FromMathMl(mathMl);
        var section = new OneNoteSection { Name = "Operator metadata" };
        var page = new OneNotePage { Title = "Math" };
        section.Pages.Add(page);
        if (entryPoint == "block-expression") {
            page.DirectContent.Add(new OneNoteMath().SetExpression(expression));
        } else if (entryPoint == "block-mathml") {
            page.DirectContent.Add(new OneNoteMath { MathMl = mathMl });
        } else {
            var paragraph = new OneNoteParagraph();
            paragraph.AddMath(entryPoint == "edited-preserved-inline"
                ? OfficeMathMarkup.FromMathMl(mathMl.Replace(" lspace='.5em'", string.Empty))
                : expression);
            page.DirectContent.Add(paragraph);
            if (entryPoint == "edited-preserved-inline") {
                section = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(section)));
                var loadedParagraph = Assert.IsType<OneNoteParagraph>(Assert.Single(
                    Assert.Single(Assert.Single(section.Pages).Outlines).Children));
                OneNoteTextRun loadedRun = Assert.Single(loadedParagraph.Runs);
                Assert.Equal(expression.ToPlainText(), loadedRun.Text);
                // Public direct replacement retains the reader's preservation state and tests
                // semantic equality at the native-reuse boundary, without a test-only hook.
                loadedRun.MathExpression = expression;
            }
        }
        byte[] sentinel = { 0x4F, 0x4E, 0x45 };
        using var destination = new MemoryStream();
        destination.Write(sentinel, 0, sentinel.Length);
        long position = destination.Position;

        OneNoteFormatException error = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(section, destination, new OneNoteWriterOptions { ValidateRoundTrip = false }));

        Assert.Equal("ONENOTE_WRITE_MATH_LOSS", error.Code);
        Assert.Equal(sentinel, destination.ToArray());
        Assert.Equal(position, destination.Position);
    }
}
