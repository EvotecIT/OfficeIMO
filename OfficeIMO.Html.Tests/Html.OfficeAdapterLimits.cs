using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlOfficeAdapterLimitTests {
    [Fact]
    public void ExcelHtml_RejectsInvalidExportRowLimit() {
        using ExcelDocument workbook = ExcelDocument.Create();
        workbook.AddWorksheet("Data").CellValue(1, 1, "value");

        Assert.Throws<ArgumentOutOfRangeException>(() => workbook.ToHtml(new ExcelHtmlSaveOptions { MaxRowsPerSheet = -1 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => workbook.ToHtml(new ExcelHtmlSaveOptions { MaxRowsPerSheet = 0 }));
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

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(new HtmlToExcelOptions { Limits = limits });
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

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
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

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
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
    public void SemanticImportModeRequiresAdapterMetadata() {
        HtmlToExcelResult result = HtmlConversionDocument.Parse("<table><tr><td>Visible only</td></tr></table>")
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Semantic });
        using ExcelDocument workbook = result.Value;

        Assert.False(result.Succeeded);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticContentMissing);
    }
}
