using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlOfficeAdapterLimitTests {
    [Fact]
    public void ExcelHtml_RejectsInvalidExportRowLimit() {
        using ExcelDocument workbook = ExcelDocument.Create();
        workbook.AddWorksheet("Data").CellValue(1, 1, "value");

        Assert.Throws<ArgumentOutOfRangeException>(() => workbook.ToHtml(new ExcelHtmlSaveOptions { MaxRowsPerSheet = -1 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => workbook.ToHtml(new ExcelHtmlSaveOptions { MaxRowsPerSheet = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => workbook.ToHtml(new ExcelHtmlSaveOptions { MaxMergedRangesPerSheet = 0 }));
    }

    [Fact]
    public void OfficeHtml_ExportsVersionedSemanticEnvelopes() {
        using ExcelDocument workbook = ExcelDocument.Create();
        workbook.AddWorksheet("Data").CellValue(1, 1, "value");
        HtmlTextConversionResult export = workbook.ToHtmlResult();
        string html = export.RequireValue();

        Assert.True(export.Succeeded);
        Assert.Contains(OfficeHtmlSemanticEnvelope.SchemaVersionAttribute + "=\"" + OfficeHtmlSemanticEnvelope.CurrentSchemaVersion + "\"", html);
    }

    [Fact]
    public void ExcelHtml_BoundsEmbeddedImageBeforeDecoding() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-profile="ExcelSemanticTables" data-officeimo-schema-version="1">
              <section class="officeimo-sheet" data-officeimo-sheet="Data" data-officeimo-range="A1">
                <table><tbody><tr><td>value</td></tr></tbody></table>
                <section class="officeimo-images"><ul><li><img src="data:image/png;base64,AQID"></li></ul></section>
              </section>
            </main>
            """;
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxImageBytes = 2;
        limits.MaxTotalImageBytes = 2;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(new HtmlToExcelOptions { Limits = limits });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(0, result.Images);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void ExcelHtml_SemanticUnsupportedImageDoesNotConsumeImageBudget() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-profile="ExcelSemanticTables" data-officeimo-schema-version="1">
              <section class="officeimo-sheet" data-officeimo-sheet="Data" data-officeimo-range="A1">
                <table><tbody><tr><td>value</td></tr></tbody></table>
                <section class="officeimo-images"><ul>
                  <li><img src="data:image/webp;base64,AA=="></li>
                  <li><img src="data:image/png;base64,AQID"></li>
                </ul></section>
              </section>
            </main>
            """;
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxImages = 1;
        limits.MaxShapes = 1;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Limits = limits });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(1, result.Images);
        Assert.Single(workbook.Sheets[0].Images);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.ResourceTypeUnsupported
                && diagnostic.Detail == "mediaType=image/webp");
    }

    [Fact]
    public void ExcelHtml_GenericUnsupportedImageDoesNotConsumeImageBudget() {
        const string html = """
            <img src="data:image/webp;base64,AA==" alt="Unsupported">
            <img src="data:image/png;base64,AQID" alt="Accepted">
            """;
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxImages = 1;
        limits.MaxShapes = 1;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Limits = limits, Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;

        ExcelImage image = Assert.Single(Assert.Single(workbook.Sheets).Images);
        Assert.Equal("Accepted", image.Description);
        Assert.Equal(1, result.Images);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.ResourceTypeUnsupported
                && diagnostic.Detail == "mediaType=image/webp");
    }

    [Fact]
    public void ExcelHtml_UsesOneFormulaDecisionForCellMetadataAndCompatibilityInventory() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-profile="ExcelSemanticTables" data-officeimo-schema-version="1">
              <section class="officeimo-sheet" data-officeimo-sheet="Data" data-officeimo-range="A1:A2">
                <table><tbody>
                  <tr><td data-officeimo-cell="A1" data-officeimo-value-kind="formula" data-officeimo-value="1+1">2</td></tr>
                  <tr><td data-officeimo-cell="A2" data-officeimo-value-kind="formula" data-officeimo-value="2+2">4</td></tr>
                </tbody></table>
                <section class="officeimo-formulas"><ul>
                  <li data-officeimo-cell="A1"><code>1+1</code></li>
                  <li data-officeimo-cell="A2"><code>2+2</code></li>
                </ul></section>
              </section>
            </main>
            """;
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxAnnotations = 1;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToExcelDocumentResult(new HtmlToExcelOptions { Limits = limits });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(1, result.Formulas);
        Assert.Single(workbook.Sheets[0].GetFormulaCells());
        Assert.Single(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
    }

    [Fact]
    public void ExcelHtml_DisabledFormulaImportKeepsVisibleFallbackText() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-profile="ExcelSemanticTables" data-officeimo-schema-version="1">
              <section class="officeimo-sheet" data-officeimo-sheet="Data" data-officeimo-range="A1">
                <table><tbody><tr><td data-officeimo-cell="A1" data-officeimo-value-kind="formula" data-officeimo-value="1+1">2</td></tr></tbody></table>
                <section class="officeimo-formulas"><ul><li data-officeimo-cell="A1"><code>1+1</code></li></ul></section>
              </section>
            </main>
            """;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(new HtmlToExcelOptions { ImportFormulas = false });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(0, result.Formulas);
        Assert.Empty(workbook.Sheets[0].GetFormulaCells());
        Assert.True(workbook.Sheets[0].TryGetCellText(1, 1, out string text));
        Assert.Equal("2", text);
    }

    [Fact]
    public void ExcelHtml_UntrustedSemanticFormulaRequiresExplicitOptIn() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-schema-version="2"
                  data-officeimo-public-semantics="safe" data-officeimo-restoration="public-safe">
              <section class="officeimo-sheet" data-officeimo-sheet="Data" data-officeimo-range="A1">
                <table><tr><td data-officeimo-cell="A1" data-officeimo-value-kind="formula" data-officeimo-value="WEBSERVICE(&quot;https://example.test&quot;)">visible</td></tr></table>
              </section>
            </main>
            """;

        HtmlToExcelResult blocked = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument blockedWorkbook = blocked.Value;

        Assert.Equal(0, blocked.Formulas);
        Assert.Empty(blockedWorkbook.Sheets[0].GetFormulaCells());
        Assert.True(blockedWorkbook.Sheets[0].TryGetCellText(1, 1, out string visibleText));
        Assert.Equal("visible", visibleText);
        Assert.Contains(blocked.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticRestorationTrustRequired);

        HtmlToExcelResult optedIn = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { AllowUntrustedFormulas = true });
        using ExcelDocument optedInWorkbook = optedIn.Value;

        Assert.Equal(1, optedIn.Formulas);
        Assert.Contains(optedInWorkbook.Sheets[0].GetFormulaCells(),
            formula => formula.CellReference == "A1" && formula.Formula.Contains("WEBSERVICE", StringComparison.Ordinal));
    }

    [Fact]
    public void ExcelHtml_GenericEmptyTableDoesNotConsumeTableOrWorksheetBudget() {
        const string html = "<table></table><table><tr><td>value</td></tr></table>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxTables = 1;
        limits.MaxSemanticContainers = 1;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Limits = limits, Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;

        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        Assert.Equal(1, result.Sheets);
        Assert.Equal(1, result.Cells);
        Assert.True(sheet.TryGetCellText(1, 1, out string text));
        Assert.Equal("value", text);
    }

    [Fact]
    public void ExcelHtml_SemanticEmptyTableDoesNotConsumeTableBudget() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-schema-version="1">
              <section class="officeimo-sheet" data-officeimo-sheet="Empty"><table></table></section>
              <section class="officeimo-sheet" data-officeimo-sheet="Data"><table><tr><td>value</td></tr></table></section>
            </main>
            """;
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxTables = 1;
        limits.MaxSemanticContainers = 2;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Limits = limits, Mode = HtmlImportMode.Semantic });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(2, result.Sheets);
        Assert.Equal(1, result.Cells);
        Assert.True(workbook.Sheets[1].TryGetCellText(1, 1, out string text));
        Assert.Equal("value", text);
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
                && diagnostic.Detail?.StartsWith(nameof(HtmlImportLimits.MaxTables), StringComparison.Ordinal) == true);
    }

    [Fact]
    public void ExcelHtml_GenericRejectedTableDoesNotConsumeNarrativeWorksheetBudget() {
        const string html = """
            <table><tr><td>first</td></tr></table>
            <table><tr><td>rejected</td></tr></table>
            <p>Narrative retained</p>
            """;
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxTables = 1;
        limits.MaxSemanticContainers = 2;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Limits = limits, Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(2, result.Sheets);
        Assert.Contains(workbook.Sheets,
            sheet => Enumerable.Range(1, 4).Any(row => sheet.CellAt(row, 1).GetValue<string>() == "Narrative retained"));
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
                && diagnostic.Detail?.StartsWith(nameof(HtmlImportLimits.MaxTables), StringComparison.Ordinal) == true);
    }

    [Fact]
    public void ExcelHtml_OmitsTextBeyondTheNativeCellLimitWithDiagnostics() {
        string html = "<main class='officeimo-document' data-officeimo-source='excel' data-officeimo-schema-version='1'>"
            + "<section class='officeimo-sheet' data-officeimo-sheet='Data'><table><tr><td>"
            + new string('x', 32_768)
            + "</td></tr></table></section></main>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument workbook = result.Value;

        Assert.Equal(0, result.Cells);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
    }

    [Fact]
    public void ExcelHtml_BoundsFormulaAndCommentCoordinatesToTheWorksheetGrid() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-schema-version="1">
              <section class="officeimo-sheet" data-officeimo-sheet="Data">
                <table><tr><td>visible</td></tr></table>
                <section class="officeimo-formulas"><ul><li data-officeimo-cell="XFE1"><code>1+1</code></li></ul></section>
                <section class="officeimo-comments"><ul><li data-officeimo-cell="A1048577"><p>outside</p></li></ul></section>
              </section>
            </main>
            """;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToExcelDocumentResult();
        using ExcelDocument workbook = result.Value;

        Assert.Equal(0, result.Formulas);
        Assert.Equal(0, result.Comments);
        Assert.Equal(2, result.Report.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticValueInvalid));
    }

    [Fact]
    public void ExcelHtml_OmitsFormulaBeyondTheNativeFormulaLimitAndKeepsVisibleText() {
        string html = "<main class='officeimo-document' data-officeimo-source='excel' data-officeimo-schema-version='1'>"
            + "<section class='officeimo-sheet' data-officeimo-sheet='Data'><table><tr>"
            + "<td data-officeimo-cell='A1' data-officeimo-value-kind='formula' data-officeimo-value='"
            + new string('1', 8_193)
            + "'>visible</td></tr></table></section></main>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToExcelDocumentResult();
        using ExcelDocument workbook = result.Value;

        Assert.Equal(0, result.Formulas);
        Assert.True(workbook.Sheets[0].TryGetCellText(1, 1, out string text));
        Assert.Equal("visible", text);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
    }

    [Fact]
    public void PowerPointHtml_BoundsInventoryChartBeforeAllocatingPlaceholderData() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="powerpoint" data-officeimo-profile="PowerPointSemanticSlides" data-officeimo-schema-version="1">
              <section class="officeimo-slide" data-officeimo-slide="1">
                <section class="officeimo-charts"><ul><li><span class="officeimo-feature-label">Huge</span><span class="officeimo-feature-meta">Series: 100000; Categories: 100000</span></li></ul></section>
              </section>
            </main>
            """;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using var presentation = result.Value;

        Assert.Equal(0, result.Charts);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void PowerPointHtml_RejectedChartDoesNotConsumeTheSharedShapeBudget() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="powerpoint" data-officeimo-profile="PowerPointSemanticSlides" data-officeimo-schema-version="1">
              <section class="officeimo-slide" data-officeimo-slide="1">
                <p data-officeimo-layer-index="1">Accepted</p>
                <section class="officeimo-charts"><ul><li data-officeimo-layer-index="0"><span class="officeimo-feature-label">Huge</span><span class="officeimo-feature-meta">Series: 100000; Categories: 100000</span></li></ul></section>
              </section>
            </main>
            """;
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxShapes = 1;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Limits = limits });
        using var presentation = result.Value;

        Assert.Equal(0, result.Charts);
        Assert.Equal(1, result.TextBoxes);
        Assert.Contains(Assert.Single(presentation.Slides).TextBoxes, textBox => textBox.Text == "Accepted");
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void PowerPointHtml_UnsupportedChartDoesNotConsumeTheSharedChartBudget() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="powerpoint" data-officeimo-profile="PowerPointSemanticSlides" data-officeimo-schema-version="1">
              <section class="officeimo-slide" data-officeimo-slide="1">
                <section class="officeimo-charts"><ul>
                  <li data-officeimo-layer-index="0"><span class="officeimo-feature-label">Unsupported</span><span class="officeimo-feature-meta">Type: FutureChart; Series: 1; Categories: 1</span></li>
                  <li data-officeimo-layer-index="1"><span class="officeimo-feature-label">Accepted</span><span class="officeimo-feature-meta">Type: Pie; Series: 1; Categories: 1</span></li>
                </ul></section>
              </section>
            </main>
            """;
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxCharts = 1;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Limits = limits });
        using var presentation = result.Value;

        Assert.Equal(1, result.Charts);
        Assert.Single(Assert.Single(presentation.Slides).Charts);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.ContentOmitted);
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void ExcelHtml_UnusableChartDoesNotConsumeTheSharedChartBudget() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-profile="ExcelSemanticTables" data-officeimo-schema-version="1">
              <section class="officeimo-sheet" data-officeimo-sheet="Data" data-officeimo-range="A1">
                <table><tbody><tr><td>value</td></tr></tbody></table>
                <section class="officeimo-charts"><ul>
                  <li data-officeimo-layer-index="0"><span class="officeimo-feature-label">Unusable</span></li>
                  <li data-officeimo-layer-index="1">
                    <span class="officeimo-feature-label">Accepted</span>
                    <table class="officeimo-chart-data">
                      <thead><tr><th></th><th>Q1</th></tr></thead>
                      <tbody><tr><th>Actual</th><td>10</td></tr></tbody>
                    </table>
                  </li>
                </ul></section>
              </section>
            </main>
            """;
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxCharts = 1;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Limits = limits });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(1, result.Charts);
        Assert.Single(Assert.Single(workbook.Sheets, sheet => sheet.Name == "Data").Charts);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.ContentOmitted);
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void WordHtml_PreparedDocumentPoliciesRemainAuthoritative() {
        const string image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
        var hyperlinkPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        hyperlinkPolicy.AllowedUrlSchemes.Clear();
        hyperlinkPolicy.AllowedUrlSchemes.Add(Uri.UriSchemeHttps);
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<p><a href='http://example.test/plain'>Plain link</a></p><img src='data:image/png;base64," + image + "' alt='Blocked image'>",
            new HtmlConversionDocumentOptions {
                Trust = HtmlInputTrust.Trusted,
                UrlPolicy = hyperlinkPolicy,
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });
        HtmlToWordOptions permissiveOptions = HtmlToWordOptions.CreateTrustedDocumentProfile();
        permissiveOptions.ResourceUrlPolicy = HtmlUrlPolicy.CreateOfficeIMOProfile();

        HtmlToWordResult result = source.ToWordDocumentResult(permissiveOptions);
        using WordDocument document = result.Value;

        Assert.Empty(document.ParagraphsHyperLinks);
        Assert.Empty(document.Images);
        Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Contains("Plain link", StringComparison.Ordinal));
        Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Contains("Blocked image", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("rules", HtmlConversionDiagnosticCodes.CssRuleLimitExceeded, nameof(HtmlConversionLimits.MaxCssRules))]
    [InlineData("declarations", HtmlConversionDiagnosticCodes.CssDeclarationLimitExceeded, nameof(HtmlConversionLimits.MaxCssDeclarations))]
    [InlineData("selectors", HtmlConversionDiagnosticCodes.CssSelectorEvaluationLimitExceeded, nameof(HtmlConversionLimits.MaxSelectorEvaluations))]
    public void WordHtml_PreparedDocumentCssComplexityLimitsRemainAuthoritative(
        string limitKind,
        string expectedCode,
        string expectedSource) {
        HtmlConversionLimits limits = HtmlConversionLimits.CreateTrustedProfile();
        string html;
        switch (limitKind) {
            case "rules":
                limits.MaxCssRules = 1;
                html = "<style>p { color:red } strong { color:blue }</style><p>text</p>";
                break;
            case "declarations":
                limits.MaxCssDeclarations = 1;
                html = "<style>p { color:red; font-weight:bold }</style><p>text</p>";
                break;
            default:
                limits.MaxSelectorEvaluations = 1;
                html = "<style>p { color:red } strong { color:blue }</style><p>text</p>";
                break;
        }

        HtmlConversionDocument source = HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Trust = HtmlInputTrust.Trusted,
            Limits = limits
        });
        HtmlToWordOptions permissiveOptions = HtmlToWordOptions.CreateTrustedDocumentProfile();

        HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
            () => source.ToWordDocumentResult(permissiveOptions));

        Assert.Equal(expectedCode, exception.Code);
        Assert.Equal(expectedSource, exception.LimitSource);
    }

    [Fact]
    public void PowerPointHtml_RejectsNonFiniteGeometryWithAVisibleFallbackDiagnostic() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="powerpoint" data-officeimo-profile="PowerPointSemanticSlides" data-officeimo-schema-version="1">
              <section class="officeimo-slide" data-officeimo-slide="1"><p data-officeimo-width="NaN">Text</p></section>
            </main>
            """;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult();
        using var presentation = result.Value;

        Assert.Equal(1, result.TextBoxes);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticValueInvalid);
    }

    [Fact]
    public void PowerPointHtml_AppliesTheSharedFieldLimitToSemanticTextBoxes() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="powerpoint" data-officeimo-schema-version="1">
              <section class="officeimo-slide"><p>oversized</p></section>
            </main>
            """;
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxMetadataCharacters = 4;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Limits = limits });
        using var presentation = result.Value;

        Assert.Equal(0, result.TextBoxes);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
    }

    [Fact]
    public void SemanticEnvelope_RejectsAnExplicitUnsupportedVersion() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-profile="ExcelSemanticTables" data-officeimo-schema-version="999">
              <section class="officeimo-sheet" data-officeimo-sheet="Data"><table><tr><td>value</td></tr></table></section>
            </main>
            """;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Auto });
        using ExcelDocument workbook = result.Value;

        Assert.False(result.Succeeded);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticSchemaUnsupported);
    }

    [Fact]
    public void SemanticEnvelope_ImportsOnlyContainersOwnedByTheValidatedRoot() {
        const string html = """
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-schema-version="1"></main>
            <main class="officeimo-document" data-officeimo-source="excel" data-officeimo-schema-version="999">
              <section class="officeimo-sheet" data-officeimo-sheet="Wrong"><table><tr><td>must not import</td></tr></table></section>
            </main>
            """;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument workbook = result.Value;

        Assert.False(result.Succeeded);
        Assert.Equal(0, result.Cells);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticContentMissing);
    }

    [Fact]
    public void ExcelHtml_AutoModeImportsOrdinaryTablesWithoutSemanticMetadata() {
        const string html = "<h1>Inventory</h1><table><caption>Products</caption><tr><th>Name</th><th>Count</th></tr><tr><td>Paper</td><td>12</td></tr></table>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Auto });
        using ExcelDocument workbook = result.Value;

        Assert.True(result.Succeeded);
        Assert.Equal(1, result.Sheets);
        Assert.Equal(4, result.Cells);
        Assert.Equal("Products", workbook.Sheets[0].Name);
    }

    [Fact]
    public void PowerPointHtml_AutoModeImportsOrdinarySectionsAndTables() {
        const string html = "<section><h2>Agenda</h2><p>First topic</p><table><tr><td>Owner</td><td>Ada</td></tr></table></section>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html).ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Auto });
        using var presentation = result.Value;

        Assert.True(result.Succeeded);
        Assert.Equal(1, result.Slides);
        Assert.Equal(2, result.TextBoxes);
        Assert.Equal(1, result.Tables);
    }

    [Fact]
    public void PowerPointHtml_GenericProjectionUsesSectionsNestedUnderMainAsSlides() {
        const string html = "<main><section><h2>First</h2><p>One</p></section><article><h2>Second</h2><p>Two</p></article></main>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using var presentation = result.Value;

        Assert.True(result.Succeeded);
        Assert.Equal(2, result.Slides);
    }

    [Fact]
    public void PowerPointHtml_GenericProjectionPreservesSiblingsAroundExplicitSections() {
        const string html = "<p>intro</p><section><h2>Middle</h2><p>inside</p></section><p>outro</p>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using var presentation = result.Value;

        Assert.True(result.Succeeded);
        Assert.Equal(3, result.Slides);
        Assert.Equal(6, result.TextBoxes);
    }

    [Fact]
    public void PowerPointHtml_GenericProjectionPreservesTextInsideOrdinaryContainers() {
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse("<div>Hello from a generic container</div>")
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using var presentation = result.Value;

        Assert.True(result.Succeeded);
        Assert.Contains(
            Assert.Single(presentation.Slides).TextBoxes,
            textBox => textBox.Text == "Hello from a generic container");
    }

    [Fact]
    public void PowerPointHtml_RejectedTextDoesNotConsumeTheSharedShapeBudget() {
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxShapes = 2;
        limits.MaxMetadataCharacters = 12;
        const string html = "<section><h2>Title</h2><p>This text is too large</p><p>Accepted</p></section>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
                Mode = HtmlImportMode.Generic,
                Limits = limits
            });
        using var presentation = result.Value;

        Assert.Equal(2, result.TextBoxes);
        Assert.Contains(Assert.Single(presentation.Slides).TextBoxes, textBox => textBox.Text == "Accepted");
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
    }

    [Fact]
    public void PowerPointHtml_RejectedPictureDoesNotConsumeTheSharedShapeBudget() {
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxShapes = 2;
        const string html = "<section><img src='https://example.test/not-embedded.png'><p>Accepted</p></section>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
                Mode = HtmlImportMode.Generic,
                Limits = limits
            });
        using var presentation = result.Value;

        Assert.Equal(0, result.Pictures);
        Assert.Equal(2, result.TextBoxes);
        Assert.Contains(Assert.Single(presentation.Slides).TextBoxes, textBox => textBox.Text == "Accepted");
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void SemanticImportModeRequiresAdapterMetadata() {
        HtmlToExcelResult result = HtmlConversionDocument.Parse("<table><tr><td>Visible only</td></tr></table>")
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Semantic });
        using ExcelDocument workbook = result.Value;

        Assert.False(result.Succeeded);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticContentMissing);
    }
}
