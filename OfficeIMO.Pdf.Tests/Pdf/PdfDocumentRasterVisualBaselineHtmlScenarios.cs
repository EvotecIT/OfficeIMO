using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
    public static IEnumerable<object[]> NativeHtmlMarketScenarioIds =>
        HtmlRenderingCorpus.All.Select(scenario => new object[] { scenario.Id });

    [Theory]
    [MemberData(nameof(NativeHtmlMarketScenarioIds))]
    public void NativeHtmlMarketScenario_MatchesPopplerRasterBaseline(string scenarioId) {
        HtmlRenderingCorpusCase scenario = HtmlRenderingCorpus.All.Single(item => item.Id == scenarioId);
        AssertScenarioRasterBaseline(
            "native-html-" + scenario.Id,
            () => CreateNativeHtmlMarketScenario(scenario),
            scenario.ExpectedPageCount);
    }

    private static byte[] CreateNativeHtmlMarketScenario(HtmlRenderingCorpusCase scenario) {
        var options = new HtmlPdfSaveOptions(scenario.CreateOptions());
        return HtmlConversionDocument.Parse(scenario.Html).ToPdf(options);
    }
}
