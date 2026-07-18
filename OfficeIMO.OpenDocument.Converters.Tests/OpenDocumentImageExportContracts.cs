using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentImageExportContracts {
    [Fact]
    public void OdtImageBridgeUsesWordRendererAndPreservesConversionLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Rendered through Word");
        source.AddTrackedParagraphInsertion("Unsupported change", "Reviewer");

        OfficeImageExportResult result = source.ExportImage(
            OfficeImageExportFormat.Png);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "ODF_IMAGE_SOURCE_TRACKED_CHANGES_UNSUPPORTED" &&
            diagnostic.LossKind == OfficeImageExportLossKind.Omission);
    }

    [Fact]
    public void OdtImageBridgeAppliesAggregateLossPolicyAfterConversion() {
        OdtDocument source = OdtDocument.Create();
        source.AddTrackedParagraphInsertion("Unsupported change", "Reviewer");
        var options = new WordImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
        };

        Assert.Throws<OfficeImageExportPolicyException>(() =>
            source.ExportImage(OfficeImageExportFormat.Png, options));
    }

    [Fact]
    public void OdsImageBridgeStreamsWorkbookSheetsThroughExcelRenderer() {
        OdsDocument source = OdsDocument.Create();
        source.AddSheet("One").Cell(0, 0).SetString("First");
        source.AddSheet("Two").Cell(0, 0).SetNumber(2);
        var names = new List<string?>();

        source.ExportImages(
            OfficeImageExportFormat.Svg,
            result => names.Add(result.Name));

        Assert.Equal(new[] { "One", "Two" }, names);
    }

    [Fact]
    public void OdpImageBridgeUsesPowerPointRendererInSlideOrder() {
        OdpPresentation source = OdpPresentation.Create();
        source.AddSlide("One").AddTextBox(
            OdfRect.FromCentimeters(1, 1, 8, 2),
            "First slide");
        source.AddSlide("Two").AddTextBox(
            OdfRect.FromCentimeters(1, 1, 8, 2),
            "Second slide");

        IReadOnlyList<OfficeImageExportResult> results =
            source.ExportImages(OfficeImageExportFormat.Png);

        Assert.Equal(2, results.Count);
        Assert.Equal("Slide 1", results[0].Name);
        Assert.Equal("Slide 2", results[1].Name);
        Assert.All(results, result => Assert.Equal(
            OfficeImageExportFormat.Png,
            result.Format));
    }
}
