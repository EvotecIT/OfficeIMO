using System.Linq;
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
    public async System.Threading.Tasks.Task DirectOpenDocumentImageSavesRetainCountsSourceAndEncoding() {
        var text = OdtDocument.Create(); text.AddParagraph("Saved ODT");
        var spreadsheet = OdsDocument.Create(); spreadsheet.AddSheet("Saved ODS").Cell(0, 0).SetString("Cell");
        var presentation = OdpPresentation.Create(); presentation.AddSlide("Saved ODP");
        string folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-odf-images-" + System.Guid.NewGuid().ToString("N"));
        try {
            OfficeConversionReport[] reports = {
                await text.SaveAsImagesAsync(System.IO.Path.Combine(folder, "odt"), new WordImageExportOptions { Format = OfficeImageExportFormat.Svg }),
                await spreadsheet.SaveAsImagesAsync(System.IO.Path.Combine(folder, "ods"), new ExcelImageExportOptions { Format = OfficeImageExportFormat.Svg }),
                await presentation.SaveAsImagesAsync(System.IO.Path.Combine(folder, "odp"), new PowerPointImageExportOptions { Format = OfficeImageExportFormat.Svg })
            };
            Assert.Equal(new[] { "ODT", "ODS", "ODP" }, reports.Select(report => report.Source));
            Assert.All(reports, report => { Assert.Equal(1, report.ItemsWritten); Assert.Equal("Svg", report.Target); });
            Assert.Equal(3, System.IO.Directory.GetFiles(folder, "*.svg", System.IO.SearchOption.AllDirectories).Length);
        } finally { if (System.IO.Directory.Exists(folder)) System.IO.Directory.Delete(folder, true); }
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void StyledOpenDocumentFamiliesExportThroughEverySharedImageFormat(OfficeImageExportFormat format) {
        OdtDocument text = OdtDocument.Create();
        OdtSpan textRun = text.AddParagraph().AddRun("Styled ODT");
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

        OfficeImageExportResult odt = Assert.Single(text.ToImages(format));
        OfficeImageExportResult ods = Assert.Single(spreadsheet.ToImages(format));
        OfficeImageExportResult odp = Assert.Single(presentation.ToImages(format));

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

        OfficeImageExportResult result = source.ToImage(
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
            source.ToImage(OfficeImageExportFormat.Png, options));
    }

    [Fact]
    public void OdsImageBridgeStreamsWorkbookSheetsThroughExcelRenderer() {
        OdsDocument source = OdsDocument.Create();
        source.AddSheet("One").Cell(0, 0).SetString("First");
        source.AddSheet("Two").Cell(0, 0).SetNumber(2);
        var names = new List<string?>();

        source.ToImages(
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
            source.ToImages(OfficeImageExportFormat.Png);

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
