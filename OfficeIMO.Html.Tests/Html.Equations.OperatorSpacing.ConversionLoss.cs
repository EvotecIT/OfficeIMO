using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlWordMathOperatorSpacingLossTests {
    [Theory]
    [InlineData(false, false, "form='prefix' lspace='.5em' rspace='0'", true)]
    [InlineData(true, true, "form='prefix' lspace='.5em' rspace='0'", true)]
    [InlineData(false, true, "form='invalid'", false)]
    [InlineData(false, true, "lspace='2ex'", true)]
    [InlineData(true, false, "rspace='calc(1em + 1px)'", true)]
    [InlineData(false, false, "", false)]
    public async Task EditableProjectionReportsAuthoredOperatorPropertiesWithoutChangingItsShape(
        bool labeled, bool asynchronous, string attributes, bool expectsOperatorLoss) {
        string label = labeled ? " aria-label='spoken fraction'" : string.Empty;
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<p>before <math" + label + "><mfrac><mrow><mi>x</mi><mo " + attributes
            + ">+</mo><mi>y</mi></mrow><mn>2</mn></mfrac></math> after</p>");
        HtmlToWordResult result = asynchronous
            ? await source.ToWordDocumentResultAsync()
            : source.ToWordDocumentResult();
        using WordDocument document = result.RequireValue();
        WordEquation equation = Assert.Single(document.Equations);
        Assert.Equal(WordEquationRepresentation.Omml, equation.Representation);
        XElement native = XElement.Parse(equation.Omml!);
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        Assert.Single(native.Descendants(math + "f"));
        Assert.Equal("x+y2", string.Concat(native.Descendants(math + "t").Select(element => element.Value)));
        Assert.Empty(document.ValidateDocument());
        HtmlDiagnostic[] losses = result.Report.Diagnostics.Where(diagnostic =>
            diagnostic.Component == "OfficeIMO.Word.Html"
            && diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated).ToArray();

        if (expectsOperatorLoss) {
            HtmlDiagnostic loss = Assert.Single(losses);
            Assert.Equal(HtmlDiagnosticSeverity.Warning, loss.Severity);
            Assert.Equal(OfficeConversionLossKind.Approximation, loss.LossKind);
            Assert.True(result.Succeeded);
            Assert.True(result.HasLoss);
            Assert.Throws<HtmlConversionException>(() => result.RequireNoLoss());
        } else {
            Assert.Empty(losses);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OperatorMetadataInsideIgnoredAnnotationIsNotAnEditableProjectionLoss(bool labeled) {
        string label = labeled ? " aria-label='spoken primary fraction'" : string.Empty;
        HtmlToWordResult result = HtmlConversionDocument.Parse(
            "<math" + label + "><semantics><mfrac><mi>a</mi><mi>b</mi></mfrac>"
            + "<annotation encoding='application/x-tex'>ignored alternate</annotation>"
            + "<annotation-xml encoding='application/mathml+xml'>"
            + "<math><mo lspace='.5em'>+</mo></math></annotation-xml></semantics></math>")
            .ToWordDocumentResult();
        using WordDocument document = result.RequireValue();

        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XElement native = XElement.Parse(Assert.Single(document.Equations).Omml!);
        Assert.Single(native.Descendants(math + "f"));
        Assert.Equal("ab", string.Concat(native.Descendants(math + "t").Select(element => element.Value)));
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Component == "OfficeIMO.Word.Html"
            && diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated);
    }
}
