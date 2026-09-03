using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSignatureProfileTests {
    [Fact]
    public void CertificationProfileEmitsDocMdpCatalogAndTransformPermissions() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Certification source"))
            .ToBytes();

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Certification,
                CertificationPermission = PdfCertificationPermissionLevel.FormFillingAndSignatures,
                FieldName = "CertificationSignature",
                ReservedSignatureContentsBytes = 512
            });
        PdfDocumentSecurityInfo security = PdfInspector.Inspect(preparation.PreparedPdf).Security;
        string raw = PdfEncoding.Latin1GetString(preparation.PreparedPdf);

        Assert.Equal(PdfSignatureProfile.Certification, preparation.Profile);
        Assert.True(security.HasDocMDPPermissions);
        Assert.Equal(2, security.DocMDPPermissionLevel);
        Assert.Contains("/Perms << /DocMDP", raw, StringComparison.Ordinal);
        Assert.Contains("/TransformMethod /DocMDP", raw, StringComparison.Ordinal);
        Assert.Contains("/P 2 /V /1.2", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void VisibleApprovalProfileCreatesWidgetAndAppearanceOnSelectedPage() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Visible approval source"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Signature target page"))
            .ToBytes();
        var appearance = new PdfVisibleSignatureAppearanceOptions {
            PageNumber = 2,
            X = 42,
            Y = 54,
            Width = 210,
            Height = 60,
            Text = "Approved by external signer"
        };

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Approval,
                FieldName = "VisibleApproval",
                VisibleAppearance = appearance,
                ReservedSignatureContentsBytes = 512
            });
        PdfReadDocument document = PdfReadDocument.Open(preparation.PreparedPdf);
        PdfFormField field = Assert.Single(document.FormFields, formField => formField.Name == "VisibleApproval");
        PdfFormWidget widget = Assert.Single(field.Widgets);
        string raw = PdfEncoding.Latin1GetString(preparation.PreparedPdf);

        Assert.Equal(2, widget.PageNumber);
        Assert.True(widget.IsPrint);
        Assert.Equal(42, widget.X1);
        Assert.Equal(54, widget.Y1);
        Assert.Equal(252, widget.X2);
        Assert.Equal(114, widget.Y2);
        Assert.Contains("Approved by external signer", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Widget", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Form", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void VisibleApprovalProfileEmbedsRasterImageInAppearanceStream() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Image-backed approval source"))
            .ToBytes();
        byte[] image = PdfPngTestImages.CreateRgbPng(4, 2);
        var appearance = new PdfVisibleSignatureAppearanceOptions {
            PageNumber = 1,
            X = 36,
            Y = 36,
            Width = 180,
            Height = 72,
            ImageBytes = image,
            ImageFit = OfficeImageFit.Contain,
            ImagePadding = 6,
            ShowText = false
        };
        image[0] = 0;

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Approval,
                FieldName = "ImageApproval",
                VisibleAppearance = appearance,
                ReservedSignatureContentsBytes = 512
            });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });
        PdfReadDocument reopened = PdfReadDocument.Open(signed);
        PdfFormWidget widget = Assert.Single(Assert.Single(reopened.FormFields, field => field.Name == "ImageApproval").Widgets);
        string raw = PdfEncoding.Latin1GetString(signed);
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(signed).Map;
        PdfStream appearanceStream = Assert.Single(objects.Values
            .Select(static item => item.Value)
            .OfType<PdfStream>(), stream =>
                PdfObjectLookup.ResolveChain(objects, stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) ? subtype : null) is PdfName { Name: "Form" } &&
                PdfObjectLookup.ResolveChain(objects, stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resources) ? resources : null) is PdfDictionary resourceDictionary &&
                resourceDictionary.Items.ContainsKey("XObject"));
        PdfDictionary appearanceResources = Assert.IsType<PdfDictionary>(
            PdfObjectLookup.ResolveChain(objects, appearanceStream.Dictionary.Items["Resources"]));
        PdfDictionary appearanceImages = Assert.IsType<PdfDictionary>(
            PdfObjectLookup.ResolveChain(objects, appearanceResources.Items["XObject"]));
        PdfStream embeddedImage = Assert.IsType<PdfStream>(
            PdfObjectLookup.ResolveChain(objects, appearanceImages.Items["Im1"]));

        Assert.Equal(1, widget.PageNumber);
        Assert.Contains("/XObject << /Im1", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Image", raw, StringComparison.Ordinal);
        Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);
        Assert.Equal("Image", Assert.IsType<PdfName>(embeddedImage.Dictionary.Items["Subtype"]).Name);
    }

    [Fact]
    public void VisibleApprovalProfileWithImageHonorsCancellationToken() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Cancelled image-backed approval source"))
            .ToBytes();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            PdfIncrementalUpdater.PrepareExternalSignature(
                source,
                new PdfExternalSignatureOptions {
                    FieldName = "CancelledImageApproval",
                    VisibleAppearance = new PdfVisibleSignatureAppearanceOptions {
                        ImageBytes = PdfPngTestImages.CreateRgbPng(4, 2),
                        ShowText = false
                    },
                    ReservedSignatureContentsBytes = 512,
                    CancellationToken = cancellation.Token
                }));
    }

    [Fact]
    public void RasterImageStreamPreparationHonorsCancellationToken() {
        byte[] imageBytes = PdfPngTestImages.CreateRgbPng(4, 2);
        var imageInfo = new OfficeImageInfo(OfficeImageFormat.Png, 4, 2);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            PdfWriter.TryBuildImageStream(
                imageBytes,
                imageInfo,
                4,
                2,
                cancellation.Token,
                out _,
                out _));
    }

    [Fact]
    public async Task PngChunkCrcValidationHonorsCancellationDuringLargeChunk() {
        const int chunkLength = 32 * 1024 * 1024;
        var png = new byte[8 + 12 + chunkLength];
        new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }.CopyTo(png, 0);
        png[8] = 2;
        png[12] = (byte)'I';
        png[13] = (byte)'D';
        png[14] = (byte)'A';
        png[15] = (byte)'T';
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(10));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Task.Run(() =>
            PdfWriter.TryGetPngImageData(
                png,
                cancellation.Token,
                out _,
                out _)));
    }

    [Fact]
    public void WideSingleRowPngTransformationHonorsCancellationInsidePackedExpansion() {
        const int width = 8193;
        byte[] png = PdfPngTestImages.CreateWidePackedGrayscalePng(width);
        var imageInfo = new OfficeImageInfo(OfficeImageFormat.Png, width, 1);
        using var cancellation = new CancellationTokenSource();
        var checkpoints = new List<int>();
        PdfWriter.PngRowLoopObserverForTesting = (kind, index) => {
            if (kind != PngRowLoopKind.PackedGrayscale) return;
            checkpoints.Add(index);
            if (index == 4096) cancellation.Cancel();
        };
        try {
            Assert.Throws<OperationCanceledException>(() =>
                PdfWriter.TryBuildImageStream(
                    png,
                    imageInfo,
                    width,
                    1,
                    cancellation.Token,
                    out _,
                    out _));
            Assert.Equal(new[] { 0, 4096 }, checkpoints);
        } finally {
            PdfWriter.PngRowLoopObserverForTesting = null;
        }
    }

    [Fact]
    public void VisibleApprovalProfileUsesAppearanceBoundsForUnidentifiedJpegDimensions() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("JPEG fallback source"))
            .ToBytes();

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                FieldName = "JpegFallback",
                VisibleAppearance = new PdfVisibleSignatureAppearanceOptions {
                    Width = 180,
                    Height = 72,
                    ImageBytes = new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 },
                    ShowText = false
                },
                ReservedSignatureContentsBytes = 512
            });
        PdfStream image = FindAppearanceImageStream(preparation.PreparedPdf);

        Assert.Equal(180D, Assert.IsType<PdfNumber>(image.Dictionary.Items["Width"]).Value);
        Assert.Equal(72D, Assert.IsType<PdfNumber>(image.Dictionary.Items["Height"]).Value);
    }

    [Theory]
    [InlineData(OfficeImageFit.Cover)]
    [InlineData(OfficeImageFit.Stretch)]
    public void VisibleApprovalProfilePaintsBorderAfterZeroPaddingImage(OfficeImageFit fit) {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Border ordering source"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                FieldName = "BorderOrdering",
                VisibleAppearance = new PdfVisibleSignatureAppearanceOptions {
                    Width = 180,
                    Height = 72,
                    ImageBytes = PdfPngTestImages.CreateRgbPng(4, 2),
                    ImageFit = fit,
                    ImagePadding = 0,
                    ShowText = false
                },
                ReservedSignatureContentsBytes = 512
            });
        PdfStream appearance = FindImageAppearanceStream(preparation.PreparedPdf);
        string content = PdfEncoding.Latin1GetString(appearance.Data);

        Assert.True(content.IndexOf("/Im1 Do", StringComparison.Ordinal) < content.LastIndexOf(" re S", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentTimestampProfileSelectsRfc3161SubFilter() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Timestamp source"))
            .ToBytes();

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.DocumentTimestamp,
                FieldName = "DocumentTimestamp",
                ReservedSignatureContentsBytes = 512
            });
        string raw = PdfEncoding.Latin1GetString(preparation.PreparedPdf);

        Assert.Equal(PdfSignatureProfile.DocumentTimestamp, preparation.Profile);
        Assert.Equal("ETSI.RFC3161", preparation.SubFilter);
        Assert.Contains("/Type /DocTimeStamp", raw, StringComparison.Ordinal);
        Assert.Contains("/SubFilter /ETSI.RFC3161", raw, StringComparison.Ordinal);
    }

    private static PdfStream FindImageAppearanceStream(byte[] pdf) {
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(pdf).Map;
        return Assert.Single(objects.Values
            .Select(static item => item.Value)
            .OfType<PdfStream>(), stream =>
                PdfObjectLookup.ResolveChain(objects, stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) ? subtype : null) is PdfName { Name: "Form" } &&
                PdfObjectLookup.ResolveChain(objects, stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resources) ? resources : null) is PdfDictionary resourceDictionary &&
                resourceDictionary.Items.ContainsKey("XObject"));
    }

    private static PdfStream FindAppearanceImageStream(byte[] pdf) {
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(pdf).Map;
        PdfStream appearance = FindImageAppearanceStream(pdf);
        PdfDictionary resources = Assert.IsType<PdfDictionary>(PdfObjectLookup.ResolveChain(objects, appearance.Dictionary.Items["Resources"]));
        PdfDictionary xObjects = Assert.IsType<PdfDictionary>(PdfObjectLookup.ResolveChain(objects, resources.Items["XObject"]));
        return Assert.IsType<PdfStream>(PdfObjectLookup.ResolveChain(objects, xObjects.Items["Im1"]));
    }
}
