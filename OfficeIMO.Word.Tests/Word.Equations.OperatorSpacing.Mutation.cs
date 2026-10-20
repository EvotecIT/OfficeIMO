using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordMathOperatorSpacingMutationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RejectedOperatorMetadataDoesNotAppendAnEquationOrParagraph(bool documentEntryPoint) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("before");
        OfficeMathExpression expression = UnsupportedExpression();

        Assert.Throws<NotSupportedException>(() => {
            if (documentEntryPoint) document.AddEquation(expression);
            else paragraph.AddEquation(expression);
        });

        Assert.Equal("before", Assert.Single(document.Paragraphs).Text);
        Assert.Empty(document.Equations);
    }

    [Fact]
    public void RejectedOperatorMetadataLeavesTheExistingEquationEditable() {
        using WordDocument document = WordDocument.Create();
        OfficeMathExpression original = OfficeMath.Fraction(OfficeMath.Identifier("a"), OfficeMath.Number("2"));
        document.AddEquation(original);
        WordEquation equation = Assert.Single(document.Equations);
        string? before = equation.Omml;

        Assert.Throws<NotSupportedException>(() => equation.SetExpression(UnsupportedExpression()));

        Assert.Equal(before, equation.Omml);
        Assert.Equal(original, Assert.Single(document.Equations).ToExpression());
        Assert.Empty(document.ValidateDocument());
    }

    private static OfficeMathExpression UnsupportedExpression() => OfficeMathMarkup.FromMathMl(
        "<math><mfrac><mrow><mi>x</mi><mo lspace='.5em'>+</mo><mi>y</mi></mrow><mn>2</mn></mfrac></math>");
}
