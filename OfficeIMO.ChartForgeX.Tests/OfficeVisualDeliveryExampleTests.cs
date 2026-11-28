#if NET8_0_OR_GREATER
using OfficeIMO.ChartForgeX.Examples;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

// Runs the shipped delivery example and reopens every saved file with OfficeIMO's own readers.
public sealed class OfficeVisualDeliveryExampleTests : IDisposable {
    private readonly string _output = Path.Combine(Path.GetTempPath(), "OfficeIMO-ChartForgeX-Delivery-" + Guid.NewGuid().ToString("N"));
    private readonly IReadOnlyList<VisualSpecimen> _specimens = VisualSpecimens.Create();

    public OfficeVisualDeliveryExampleTests() {
        Directory.CreateDirectory(_output);
        DocumentDelivery.Write(_output, _specimens);
    }

    public void Dispose() {
        if (Directory.Exists(_output)) Directory.Delete(_output, recursive: true);
    }

    [Fact]
    public void ExampleCoversChartTopologyAndDiagramInBothThemes() {
        Assert.Equal(new[] { "capacity-light", "delivery-light", "approval-light", "capacity-dark", "delivery-dark", "approval-dark",
            "capacity-light-repeat" }, _specimens.Select(item => item.Name));
        Assert.All(_specimens, item => Assert.False(string.IsNullOrWhiteSpace(item.Artifact.Accessibility.Description)));
        // Every specimen keeps a vector placement payload; PNG remains the reported fallback for incomplete SVG import.
        Assert.All(_specimens, item => Assert.Equal(OfficeVisualMediaFormat.Svg, item.Convert(450D).PlacementFormat));
    }

    [Fact]
    public void WordPackageHoldsEveryVisualAndValidates() {
        using WordDocument document = WordDocument.Load(Path.Combine(_output, "service-review.docx"));
        Assert.Empty(document.ValidateDocument());
        Assert.Equal(_specimens.Count, document.Images.Count);
        foreach (VisualSpecimen specimen in _specimens)
            Assert.Contains(document.Images, image => image.AltText == specimen.Caption);
    }

    [Fact]
    public void ExcelPackagePlacesOneVisualPerSheetAndValidates() {
        using ExcelDocument workbook = ExcelDocument.Load(Path.Combine(_output, "service-review.xlsx"));
        Assert.Empty(workbook.ValidateDocument());
        Assert.Equal(_specimens.Select(item => item.Name), workbook.Sheets.Select(sheet => sheet.Name));
        Assert.All(workbook.Sheets, sheet => Assert.Single(sheet.Images));
    }

    [Fact]
    public void PowerPointPackagePlacesOneVisualPerSlideAndValidates() {
        using PowerPointDocument presentation = PowerPointDocument.Load(Path.Combine(_output, "service-review.pptx"));
        Assert.Empty(presentation.ValidateDocument());
        Assert.Equal(_specimens.Count, presentation.Slides.Count);
        for (int index = 0; index < _specimens.Count; index++) {
            PowerPointImage picture = Assert.Single(presentation.Slides[index].Images);
            Assert.Equal(_specimens[index].Caption, picture.AltText);
        }
    }

    [Fact]
    public void PdfKeepsEachVisualAsVectorTextOnItsOwnPage() {
        using PdfDocument pdf = PdfDocument.Load(Path.Combine(_output, "service-review.pdf"));
        Assert.Equal(_specimens.Count, pdf.Pages.Count);
        Assert.Empty(pdf.Images.Extract());
        string[] labels = { "Capacity", "Processing worker", "Review risk" };
        for (int index = 0; index < _specimens.Count; index++) {
            string text = pdf.Pages[index].ExtractText();
            Assert.Contains(_specimens[index].Artifact.Title, text);
            Assert.Contains(labels[index % 3], text);
        }
    }

    [Fact]
    public void VisioBookKeepsEditableTopologyAndFlowPages() {
        string path = Path.Combine(_output, "service-delivery.vsdx");
        Assert.Empty(VisioValidator.Validate(path));
        VisioDocument document = VisioDocument.Load(path);
        Assert.Equal(4, document.Pages.Count);
        foreach (VisioPage page in document.Pages.Where((_, index) => index % 2 == 0)) {
            Assert.Equal(new[] { "gateway", "worker", "ledger" }, page.Shapes.Where(shape => shape.GetShapeDataValue("CFX.TopologyNodeKind") != null)
                .Select(shape => shape.Id));
            Assert.Equal(new[] { "accept", "persist" }, page.Connectors.Select(connector => connector.Id));
        }
        foreach (VisioPage page in document.Pages.Where((_, index) => index % 2 == 1)) {
            Assert.Equal("Decision", page.Shapes.Single(shape => shape.Id == "review").GetShapeDataValue("CFX.FlowStepKind"));
            Assert.Equal("Positive", page.Shapes.Single(shape => shape.Id == "deploy").GetShapeDataValue("CFX.Status"));
            Assert.Equal(4, page.Connectors.Count);
            Assert.Contains(page.Connectors, connector => connector.Label == "Rejected");
        }
    }
}
#endif
