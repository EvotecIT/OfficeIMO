using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingMathTextTransformTests {
    [Fact]
    public void MathematicalAutoTextCoversLatinGreekAndExceptionalMappings() {
        // MathML Core appendix C.1: endpoints, the Greek hole, and non-contiguous aliases.
        const string source = "AhZazıȷΑΡϴΣΩαω∂∇ϵϑϰϕϱϖ";
        const string expected = "𝐴ℎ𝑍𝑎𝑧𝚤𝚥𝛢𝛲𝛳𝛴𝛺𝛼𝜔𝜕𝛻𝜖𝜗𝜘𝜙𝜚𝜛";
        Assert.Equal(expected, string.Concat(source.Select(c => OfficeMathTextTransform.MathAuto(c.ToString()))));
        foreach (string text in new[] { "", "xy", "x́", "𝑥", "∞", "\u03A2", "\ud800" })
            Assert.Equal(text, OfficeMathTextTransform.MathAuto(text));
    }
    [Fact]
    public void MathMlNestedRowsParseEachSourceTokenOnce() {
        string markup = "<mi>x</mi>";
        for (int i = 0; i < 6; i++) markup = "<mrow>" + markup + "<mi>y</mi></mrow>";
        var source = System.Xml.Linq.XElement.Parse(markup);
        var visited = new List<System.Xml.Linq.XElement>();
        OfficeMathExpression expression = OfficeMathMarkup.FromMathMl(source, (token, _) => visited.Add(token));
        Assert.Equal("xyyyyyy", expression.ToPlainText());
        Assert.Equal(source.Descendants().Count(e => e.Name.LocalName == "mi"), visited.Count);
        Assert.Equal(visited.Count, visited.Distinct().Count());
    }

}
