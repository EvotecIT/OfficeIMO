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
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void StyledOpenDocumentFamiliesExportThroughEverySharedImageFormat(OfficeImageExportFormat format) {
        OdtDocument text = OdtDocument.Create();
        OdtSpan textRun = text.AddParagraph().AddSpan("Styled ODT");
        textRun.Bold = true;
        textRun.Italic = true;
        textRun.UnderlineStyle = OdfTextDecorationStyle.Wave;
        textRun.TextPosition = OdfTextPosition.Superscript;

        OdsDocument spreadsheet = OdsDocument.Create();
        OdsCell cell = spreadsheet.AddSheet("Typography").Cell(0, 0);
        cell.SetString("Styled ODS");
        cell.Bold = true;
        cell.Italic = true;
        cell.UnderlineStyle = OdfTextDecorationStyle.Dotted;
        cell.TextPosition = OdfTextPosition.Subscript;

        OdpPresentation presentation = OdpPresentation.Create();
        OdpRun slideRun = presentation.AddSlide("Typography")
            .AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), null, "Text")
            .AddParagraph().AddRun("Styled ODP");
        slideRun.Bold = true;
        slideRun.Italic = true;
        slideRun.UnderlineStyle = OdfTextDecorationStyle.Wave;
        slideRun.TextPosition = OdfTextPosition.Superscript;

        OfficeImageExportResult odt = Assert.Single(text.ExportImages(format));
        OfficeImageExportResult ods = Assert.Single(spreadsheet.ExportImages(format));
        OfficeImageExportResult odp = Assert.Single(presentation.ExportImages(format));

        Assert.All(new[] { odt, ods, odp }, result => {
            Assert.Equal(format, result.Format);
            Assert.True(result.Bytes.Length > 32);
        });
    }

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
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
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

    [Fact]
    public void ConversionDiagnosticsPreserveBatchSequenceMetadata() {
        OfficeImageExportResult? sequenced = null;
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        OfficeImageExportBatchProcessor.Run(
            new OfficeImageExportOptions(),
            (accept, _) => accept(new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png)),
            result => sequenced = result,
            expectedOutputCount: 1);
        var report = new OdfConversionReport("ODT", "Word")
            .Add("tracked-changes", OdfConversionMappingStatus.Unsupported);

        OfficeImageExportResult attached = OdfImageExportDiagnostics.Attach(sequenced!, report);

        Assert.Equal(0, attached.SequenceIndex);
        Assert.Equal(1, attached.SequenceCount);
        Assert.Single(attached.Diagnostics);
    }
}
