using System.IO.Compression;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserNotebookConversionTests {
    private static readonly ConversionRoute NotebookPdf = ConversionRouteCatalog.Find("notebook-pdf");
    private readonly BrowserConversionService _service = new();

    [Fact]
    public void DefaultSupportBundleOmitsTheNotebookTextInMathWarnings() {
        const string privateMarker = "PrivateLabCellMarker";
        SelectedDocument notebook = Notebook("""
            {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
              {"id":"results","cell_type":"code","metadata":{},"execution_count":1,"source":"","outputs":[
                {"output_type":"display_data","metadata":{},"data":{"text/latex":"\\begin{PrivateLabCellMarker}1\\end{PrivateLabCellMarker}"}}]}
            ]}
            """u8.ToArray());

        ConversionResult result = _service.ConvertFile(NotebookPdf, notebook, limitExcelRows: false);
        Assert.Contains(result.StructuredWarnings, warning => warning.Code == "math-unsupported"
            && warning.Source.Contains(privateMarker, StringComparison.Ordinal));

        BrowserConversionArtifact bundle = _service.CreateSupportBundle(result);
        using var archive = new ZipArchive(new MemoryStream(bundle.Bytes), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry("support-summary.json")!.Open());
        string summaryText = reader.ReadToEnd();
        using JsonDocument summary = JsonDocument.Parse(summaryText);
        Assert.Contains(summary.RootElement.GetProperty("warnings").EnumerateArray(),
            warning => warning.GetProperty("Code").GetString() == "math-unsupported");
        Assert.DoesNotContain(privateMarker, summaryText, StringComparison.Ordinal);
        Assert.Contains(summary.RootElement.GetProperty("engine").GetProperty("supplementalFonts").EnumerateArray(),
            asset => asset.GetProperty("family").GetString() == "STIX Two Math");
    }

    [Fact]
    public void NotebookManifestIdentifiesPresentationSettingsAndPortableFonts() {
        const string customTitle = "Private report title";
        SelectedDocument notebook = Notebook("""
            {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
              {"id":"results","cell_type":"code","metadata":{},"execution_count":1,"source":"visible_code()","outputs":[
                {"output_type":"stream","name":"stdout","text":"Saved measurement"}
              ]}
            ]}
            """u8.ToArray());
        var settings = new BrowserNotebookSettings(true, false, customTitle, string.Empty);

        ConversionResult baseline = _service.ConvertFile(NotebookPdf, notebook, limitExcelRows: false);
        ConversionResult changed = ConvertWith(settings);
        ConversionResult repeated = ConvertWith(settings);

        string text = PdfReadDocument.Open(changed.Bytes).ExtractText();
        Assert.Contains("Saved measurement", text, StringComparison.Ordinal);
        Assert.DoesNotContain("visible_code", text, StringComparison.Ordinal);
        using JsonDocument before = JsonDocument.Parse(baseline.CompanionReport!.Bytes);
        using JsonDocument after = JsonDocument.Parse(changed.CompanionReport!.Bytes);
        using JsonDocument again = JsonDocument.Parse(repeated.CompanionReport!.Bytes);
        Assert.NotEqual(before.RootElement.GetProperty("conversionId").GetString(),
            after.RootElement.GetProperty("conversionId").GetString());
        Assert.Equal(after.RootElement.GetProperty("conversionId").GetString(),
            again.RootElement.GetProperty("conversionId").GetString());
        Assert.DoesNotContain(customTitle, after.RootElement.GetRawText(), StringComparison.Ordinal);
        JsonElement fonts = after.RootElement.GetProperty("fontPack").GetProperty("supplementalFonts");
        Assert.Contains(fonts.EnumerateArray(), asset => asset.GetProperty("family").GetString() == "STIX Two Math"
            && asset.GetProperty("sha256").GetString() == "49b22767985171bce579307b21d98f4d35012fa39e316acc4ffcab8644dfb69b");
        Assert.Contains(fonts.EnumerateArray(), asset => asset.GetProperty("family").GetString() == "Liberation Mono"
            && asset.GetProperty("sha256").GetString() == "f2b83c763e8afd21709333370bed4774337fae82267937e2b5aea7e2fbd922c1");

        ConversionResult ConvertWith(BrowserNotebookSettings options) => _service.ConvertUpload(
            NotebookPdf, notebook, limitExcelRows: false, profile: null, generateDebugOverlay: false,
            powerPointImportProfileId: null, notebookSettings: options);
    }

    private static SelectedDocument Notebook(byte[] bytes) => new("lab.ipynb", ".ipynb", "IPYNB", bytes.LongLength, bytes);
}
