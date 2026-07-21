using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void SemanticEnvelopeV2_EmitsPublicSafeContractAndAcceptsV1() {
        var builder = new StringBuilder("<main class=\"officeimo-document\"");
        OfficeHtmlSemanticEnvelope.AppendRootAttributes(builder, "excel", "SemanticTables");
        builder.Append("></main>");
        string v2 = builder.ToString();

        Assert.Contains("data-officeimo-schema-version=\"2\"", v2, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-public-semantics=\"safe\"", v2, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-restoration=\"public-safe\"", v2, StringComparison.Ordinal);

        var v2Document = HtmlDocumentParser.ParseDocument(v2);
        OfficeHtmlSemanticEnvelopeInfo v2Info = OfficeHtmlSemanticEnvelope.Inspect(v2Document, "excel");
        Assert.True(v2Info.IsSupported);
        Assert.True(v2Info.CanRestoreTargetSpecific(HtmlInputTrust.Untrusted));

        var v1Document = HtmlDocumentParser.ParseDocument(
            "<main class='officeimo-document' data-officeimo-source='excel' data-officeimo-schema-version='1'></main>");
        OfficeHtmlSemanticEnvelopeInfo v1Info = OfficeHtmlSemanticEnvelope.Inspect(v1Document, "excel");
        Assert.True(v1Info.IsSupported);
        Assert.False(v1Info.IsLegacy);
    }

    [Fact]
    public void SemanticEnvelopeV2_RequiresExplicitTrustForTrustedTargetRestorationWithoutSilentLoss() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel"
                  data-officeimo-schema-version="2" data-officeimo-public-semantics="safe"
                  data-officeimo-restoration="trusted-target">
              <section class="officeimo-sheet" data-officeimo-sheet="Data" data-officeimo-range="A1:B2">
                <table><tr><th>Name</th><th>Value</th></tr><tr><td>Total</td><td>42</td></tr></table>
              </section>
            </main>
            """;

        HtmlToExcelResult untrusted = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using OfficeIMO.Excel.ExcelDocument publicSafe = untrusted.RequireValue();
        Assert.Contains(untrusted.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticRestorationTrustRequired
                && diagnostic.LossKind == HtmlConversionLossKind.Approximation);

        HtmlConversionDocument trustedSource = HtmlConversionDocument.Parse(
            html,
            HtmlConversionDocumentOptions.CreateTrustedProfile());
        using OfficeIMO.Excel.ExcelDocument trusted = trustedSource.ToExcelDocumentResult().RequireValue();
        Assert.Equal("Data", Assert.Single(trusted.Sheets).Name);
    }

    [Theory]
    [InlineData("", "public-safe")]
    [InlineData("unknown", "public-safe")]
    [InlineData("safe", "")]
    [InlineData("safe", "future-mode")]
    public void SemanticEnvelopeV2_FailsClosedForMissingOrUnknownSafetyContracts(string publicSemantics, string restoration) {
        string html = "<main class='officeimo-document' data-officeimo-source='excel'"
            + " data-officeimo-schema-version='2' data-officeimo-public-semantics='" + publicSemantics + "'"
            + " data-officeimo-restoration='" + restoration + "'>"
            + "<section class='officeimo-sheet' data-officeimo-sheet='Unsafe'><table><tr><td data-officeimo-formula='=1+1'>2</td></tr></table></section>"
            + "</main>";

        OfficeHtmlSemanticEnvelopeInfo envelope = OfficeHtmlSemanticEnvelope.Inspect(
            HtmlDocumentParser.ParseDocument(html), "excel");
        Assert.False(envelope.IsSupported);
        Assert.False(envelope.CanRestoreTargetSpecific(HtmlInputTrust.Trusted));

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        Assert.False(result.Succeeded);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticSchemaUnsupported
                && diagnostic.LossKind == HtmlConversionLossKind.Failure);
        result.Value.Dispose();
    }
}
