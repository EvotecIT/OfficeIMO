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
    public void VisibleApprovalProfileEmbedsOrientedJpegInAppearanceStream() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Image-backed approval source"))
            .ToBytes();
        var raster = new OfficeRasterImage(2, 1);
        raster.SetPixel(0, 0, OfficeColor.Red);
        raster.SetPixel(1, 0, OfficeColor.Blue);
        byte[] image = OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions {
            Quality = 100,
            Subsampling = OfficeJpegSubsampling.Y444,
            Metadata = new OfficeJpegMetadata(exif: [
                (byte)'I', (byte)'I', 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
                0x01, 0x00, 0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00,
                0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
            ])
        });
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
        Assert.Equal(1, Assert.IsType<PdfNumber>(embeddedImage.Dictionary.Items["Width"]).Value);
        Assert.Equal(2, Assert.IsType<PdfNumber>(embeddedImage.Dictionary.Items["Height"]).Value);
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
    public void PngChunkCrcValidationHonorsCancellationDuringLargeChunk() {
        const int chunkLength = 8192;
        var png = new byte[8 + 12 + chunkLength];
        new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }.CopyTo(png, 0);
        png[10] = (byte)(chunkLength >> 8);
        png[11] = (byte)(chunkLength & 0xFF);
        png[12] = (byte)'I';
        png[13] = (byte)'D';
        png[14] = (byte)'A';
        png[15] = (byte)'T';
        using var cancellation = new CancellationTokenSource();
        var checkpoints = new List<int>();
        PdfWriter.PngRowLoopObserverForTesting = (kind, index) => {
            if (kind != PngRowLoopKind.CrcValidation) return;
            checkpoints.Add(index);
            if (index == 4096) cancellation.Cancel();
        };
        try {
            Assert.Throws<OperationCanceledException>(() =>
                PdfWriter.TryGetPngImageData(
                    png,
                    cancellation.Token,
                    out _,
                    out _));
            Assert.Equal(new[] { 0, 4096 }, checkpoints);
        } finally {
            PdfWriter.PngRowLoopObserverForTesting = null;
        }
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

    [Fact]
    public void VisibleApprovalProfileOmitsBackgroundAndBorderWhenDisabled() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Image only appearance source"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                FieldName = "ImageOnly",
                VisibleAppearance = new PdfVisibleSignatureAppearanceOptions {
                    Width = 180,
                    Height = 72,
                    ImageBytes = PdfPngTestImages.CreateRgbPng(4, 2),
                    ImagePadding = 0,
                    ShowText = false,
                    ShowBackground = false,
                    ShowBorder = false
                },
                ReservedSignatureContentsBytes = 512
            });
        PdfStream appearance = FindImageAppearanceStream(preparation.PreparedPdf);
        string content = PdfEncoding.Latin1GetString(appearance.Data);

        Assert.Contains("/Im1 Do", content, StringComparison.Ordinal);
        Assert.DoesNotContain(" re f", content, StringComparison.Ordinal);
        Assert.DoesNotContain(" re S", content, StringComparison.Ordinal);
        Assert.DoesNotContain("BT ", content, StringComparison.Ordinal);
    }

    [Fact]
    public void ApprovalProfileAllowsAdditionalSignatureOverSignedRevision() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Multi-signature source"))
            .ToBytes();
        PdfExternalSignaturePreparation first = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions { FieldName = "First", ReservedSignatureContentsBytes = 512 });
        byte[] firstSigned = PdfIncrementalUpdater.ApplyExternalSignature(first, new byte[] { 0x30, 0x03, 0x02, 0x01, 0x01 });

        PdfAppendOnlyMutationReport mutation = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(firstSigned);
        PdfExternalSignaturePreparation second = PdfIncrementalUpdater.PrepareExternalSignature(
            firstSigned,
            new PdfExternalSignatureOptions { FieldName = "Second", ReservedSignatureContentsBytes = 512 });
        byte[] secondSigned = PdfIncrementalUpdater.ApplyExternalSignature(second, new byte[] { 0x30, 0x03, 0x02, 0x01, 0x02 });
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(secondSigned);

        Assert.True(mutation.CanPrepareExternalSignature);
        Assert.Contains("SignedAdditionalSignature", mutation.Warnings);
        Assert.True(secondSigned.AsSpan(0, firstSigned.Length).SequenceEqual(firstSigned));
        Assert.Equal(2, report.Signatures.Count);
        Assert.True(report.IsStructurallyValid);
        Assert.Contains(report.Signatures, signature => signature.Signature.FieldName == "First");
        Assert.Contains(report.Signatures, signature => signature.Signature.FieldName == "Second");
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    public void CertificationWithPermittedChangesAllowsAdditionalSignaturePreparation(int permissionLevel) {
        byte[] certified = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(permissionLevel);

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            certified,
            new PdfExternalSignatureOptions { FieldName = "Second", ReservedSignatureContentsBytes = 512 });
        PdfDocumentSecurityInfo security = PdfInspector.Inspect(preparation.PreparedPdf).Security;

        Assert.True(preparation.PreparedPdf.AsSpan(0, certified.Length).SequenceEqual(certified));
        Assert.Equal(permissionLevel, security.DocMDPPermissionLevel);
        Assert.Equal(2, security.SignatureFieldCount);
    }

    [Fact]
    public void AdditionalSignaturePreparationBlocksAFieldIncludedBySignatureLock() {
        byte[] certified = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(
            permissionLevel: 2,
            lockDictionary: "<< /Type /SigFieldLock /Action /Include /Fields [(Second)] >>");
        var options = new PdfExternalSignatureOptions { FieldName = "Second", ReservedSignatureContentsBytes = 512 };

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfIncrementalUpdater.PrepareExternalSignature(certified, options));
        PdfOperationResult<PdfExternalSignaturePreparation> result = PdfDocument.Load(certified)
            .PrepareExternalSignatureResult(options);

        Assert.Contains("AppendOnly.SignatureFieldLock", exception.Plan.BlockerCodes);
        Assert.False(result.Succeeded);
        Assert.Contains("AppendOnly.SignatureFieldLock", Assert.IsType<PdfMutationPlan>(result.MutationPlan).BlockerCodes);
    }

    [Fact]
    public void AdditionalSignaturePreparationAllowsAFieldOutsideSignatureLock() {
        byte[] certified = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(
            permissionLevel: 2,
            lockDictionary: "<< /Type /SigFieldLock /Action /Include /Fields [(Name)] >>");

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            certified,
            new PdfExternalSignatureOptions { FieldName = "Second", ReservedSignatureContentsBytes = 512 });

        Assert.True(preparation.PreparedPdf.AsSpan(0, certified.Length).SequenceEqual(certified));
        Assert.Equal("Second", preparation.FieldName);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CertificationProfileRequiresTheFirstSignedRevision(bool existingCertification) {
        byte[] signed;
        if (existingCertification) {
            signed = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(permissionLevel: 2);
        } else {
            byte[] source = PdfDocument.Create()
                .Paragraph(paragraph => paragraph.Text("Approval source"))
                .ToBytes();
            PdfExternalSignaturePreparation approval = PdfIncrementalUpdater.PrepareExternalSignature(
                source,
                new PdfExternalSignatureOptions { FieldName = "Approval", ReservedSignatureContentsBytes = 512 });
            signed = PdfIncrementalUpdater.ApplyExternalSignature(approval, new byte[] { 0x30, 0x03, 0x02, 0x01, 0x01 });
        }
        var options = new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            FieldName = "Certification",
            ReservedSignatureContentsBytes = 512
        };

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfIncrementalUpdater.PrepareExternalSignature(signed, options));
        PdfOperationResult<PdfExternalSignaturePreparation> result = PdfDocument.Load(signed)
            .PrepareExternalSignatureResult(options);

        Assert.Contains("AppendOnly.CertificationRequiresFirstSignature", exception.Plan.BlockerCodes);
        Assert.False(result.Succeeded);
        Assert.Contains(
            "AppendOnly.CertificationRequiresFirstSignature",
            Assert.IsType<PdfMutationPlan>(result.MutationPlan).BlockerCodes);
    }

    [Fact]
    public void CertificationProfileAllowsTheFirstSignatureWhenOnlyAnUnsignedFieldExists() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Unsigned field source"))
            .ToBytes();
        byte[] unsigned = PdfAcroFormEditor.Edit(
            source,
            edit => edit.PlaceSignatureField("EmptySignature", 1, 72, 500, 180, 40)).ToBytes();
        unsigned = PdfDocumentObjectGraphRewriter.Rewrite(unsigned, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[Assert.IsType<int>(security.RootObjectNumber)].Value);
            PdfDictionary acroForm = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["AcroForm"]));
            acroForm.Items["SigFlags"] = new PdfNumber(1);
            return security.InfoObjectNumber;
        });
        PdfDocumentSecurityInfo security = PdfInspector.Inspect(unsigned).Security;
        var options = new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            FieldName = "Certification",
            ReservedSignatureContentsBytes = 512
        };

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(unsigned, options);
        PdfOperationResult<PdfExternalSignaturePreparation> result = PdfDocument.Load(unsigned)
            .PrepareExternalSignatureResult(options);

        Assert.True(security.AcroFormSignaturesExist);
        Assert.Equal(0, security.SignatureValueCount);
        Assert.Equal(PdfSignatureProfile.Certification, preparation.Profile);
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void SignedFieldMdpBlocksAFieldWithoutMutableSignatureLock() {
        byte[] certified = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(
            permissionLevel: 2,
            fieldMdpTransformParameters: "<< /Type /TransformParams /V /1.2 /Action /Include /Fields [(Second)] >>");
        var options = new PdfExternalSignatureOptions { FieldName = "Second", ReservedSignatureContentsBytes = 512 };

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfIncrementalUpdater.PrepareExternalSignature(certified, options));
        PdfSignatureFieldLockInfo fieldLock = Assert.IsType<PdfSignatureFieldLockInfo>(
            Assert.Single(PdfInspector.Inspect(certified).Security.Signatures).FieldLock);

        Assert.Contains("AppendOnly.SignatureFieldLock", exception.Plan.BlockerCodes);
        Assert.True(fieldLock.LocksIncludedFields);
        Assert.Equal(new[] { "Second" }, fieldLock.Fields);
    }

    [Fact]
    public void SignedFieldMdpOverridesDivergentMutableSignatureLock() {
        byte[] certified = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(
            permissionLevel: 2,
            lockDictionary: "<< /Type /SigFieldLock /Action /Include /Fields [(Name)] >>",
            fieldMdpTransformParameters: "<< /Type /TransformParams /V /1.2 /Action /Include /Fields [(Second)] >>");

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfIncrementalUpdater.PrepareExternalSignature(
                certified,
                new PdfExternalSignatureOptions { FieldName = "Second", ReservedSignatureContentsBytes = 512 }));

        Assert.Contains("AppendOnly.SignatureFieldLock", exception.Plan.BlockerCodes);
    }

    [Fact]
    public void SignedFieldMdpAllowsAFieldDespiteDivergentMutableSignatureLock() {
        byte[] certified = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(
            permissionLevel: 2,
            lockDictionary: "<< /Type /SigFieldLock /Action /Include /Fields [(Second)] >>",
            fieldMdpTransformParameters: "<< /Type /TransformParams /V /1.2 /Action /Include /Fields [(Name)] >>");

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            certified,
            new PdfExternalSignatureOptions { FieldName = "Second", ReservedSignatureContentsBytes = 512 });

        Assert.True(preparation.PreparedPdf.AsSpan(0, certified.Length).SequenceEqual(certified));
        Assert.Equal("Second", preparation.FieldName);
    }

    [Theory]
    [InlineData("<< /TransformMethod /FieldMDP /TransformParams << /Action /Include /Fields [(Second)] >> >>")]
    [InlineData("[10 0 R]")]
    [InlineData("[<< /TransformMethod /FieldMDP /TransformParams << /Action /Include /Fields [(Second) 42] >> >>]")]
    public void MalformedSignedFieldMdpFailsClosed(string signatureReference) {
        byte[] signed = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(
            permissionLevel: 2,
            signatureReference: signatureReference,
            includeDocMdpPermissions: false);
        var options = new PdfExternalSignatureOptions { FieldName = "Second", ReservedSignatureContentsBytes = 512 };

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfIncrementalUpdater.PrepareExternalSignature(signed, options));
        PdfSignatureFieldLockInfo fieldLock = Assert.IsType<PdfSignatureFieldLockInfo>(
            Assert.Single(PdfInspector.Inspect(signed).Security.Signatures).FieldLock);

        Assert.True(fieldLock.LocksAllFields);
        Assert.Contains("AppendOnly.SignatureFieldLock", exception.Plan.BlockerCodes);
    }

    [Fact]
    public void CertificationWithNoChangesBlocksAdditionalSignaturePreparation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Locked certification source"))
            .ToBytes();
        PdfExternalSignaturePreparation certification = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Certification,
                CertificationPermission = PdfCertificationPermissionLevel.NoChanges,
                FieldName = "Certification",
                ReservedSignatureContentsBytes = 512
            });
        byte[] certified = PdfIncrementalUpdater.ApplyExternalSignature(certification, new byte[] { 0x30, 0x03, 0x02, 0x01, 0x01 });

        PdfAppendOnlyMutationReport mutation = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(certified);

        Assert.False(mutation.CanPrepareExternalSignature);
        Assert.Contains("SignaturePrepare", mutation.BlockedActions);
        Assert.Throws<PdfMutationBlockedException>(() => PdfIncrementalUpdater.PrepareExternalSignature(
            certified,
            new PdfExternalSignatureOptions { FieldName = "Second", ReservedSignatureContentsBytes = 512 }));
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
