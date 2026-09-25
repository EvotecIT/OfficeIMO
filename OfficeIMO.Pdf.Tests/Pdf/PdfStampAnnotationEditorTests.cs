using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfStampAnnotationEditorTests {
    [Fact]
    public void ImageStampIsAReopenableAnnotationWithAnImageAppearance() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Image stamp source")).ToBytes();
        byte[] image = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(24, 12, OfficeColor.Black), OfficeImageExportFormat.Png);
        var options = new PdfStampAnnotationOptions {
            StampName = "Signature", ImageBytes = image, X = 72, Y = 600, Width = 120, Height = 60,
            Contents = "Visual signature"
        };
        image[0] = 0; // The option owns its image bytes after assignment.

        PdfAnnotationEditResult result = PdfDocument.Load(source).Annotations.AddStamp(options);

        PdfAnnotation stamp = Assert.Single(PdfDocument.Load(result.Bytes).Reader.Annotations());
        Assert.Equal("Stamp", stamp.Subtype);
        Assert.True(stamp.HasNormalAppearance);
        Assert.Equal("Visual signature", stamp.Contents);
        Assert.Equal(72, stamp.X1);
        Assert.Equal(600, stamp.Y1);
        Assert.Equal(192, stamp.X2);
        Assert.Equal(660, stamp.Y2);
        var (objects, _) = PdfSyntax.ParseObjects(result.Bytes);
        PdfStream appearance = Assert.Single(objects.Values.Select(static item => item.Value).OfType<PdfStream>(),
            static stream => stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) && subtype is PdfName name && name.Name == "Form");
        Assert.Contains("/Im1 Do", PdfEncoding.Latin1GetString(appearance.Data), StringComparison.Ordinal);
        Assert.Contains(objects.Values.Select(static item => item.Value).OfType<PdfStream>(),
            static stream => stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) && subtype is PdfName name && name.Name == "Image");
        byte[] beforePng = PdfPageImageRenderer.RenderPageAsPng(source, scale: 0.5D);
        byte[] afterPng = PdfPageImageRenderer.RenderPageAsPng(result.Bytes, scale: 0.5D);
        Assert.NotEqual(beforePng, afterPng);
    }

    [Fact]
    public void ImageStampCanBeAppendedToAnnotationPermittedCertification() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Signed source")).ToBytes();
        byte[] certified = Certify(source, PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures);
        byte[] image = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(4, 4, OfficeColor.Black), OfficeImageExportFormat.Png);

        PdfAnnotationEditResult result = PdfDocument.Load(certified).Annotations.AddStamp(
            new PdfStampAnnotationOptions { StampName = "Signature", ImageBytes = image });

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.Bytes.AsSpan(0, certified.Length).SequenceEqual(certified));
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.True(Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Stamp")).HasNormalAppearance);
    }

    [Fact]
    public void TransparentImageStampEmbedsItsSoftMask() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Signature area")).ToBytes();
        var raster = new OfficeRasterImage(4, 4, OfficeColor.Transparent);
        raster.SetPixel(1, 1, OfficeColor.Black);
        byte[] image = OfficeRasterImageEncoder.Encode(raster, OfficeImageExportFormat.Png);

        PdfAnnotationEditResult result = PdfDocument.Load(source).Annotations.AddStamp(
            new PdfStampAnnotationOptions { ImageBytes = image, Width = 80, Height = 80 });

        var (objects, _) = PdfSyntax.ParseObjects(result.Bytes);
        Assert.Contains(objects.Values.Select(static item => item.Value).OfType<PdfStream>(),
            static stream => stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) &&
                             subtype is PdfName { Name: "Image" } && stream.Dictionary.Items.ContainsKey("SMask"));
    }

    [Fact]
    public void ImageStampRejectsOversizedInputBeforeProducingAnArtifact() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Source")).ToBytes();
        byte[] image = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(4, 4, OfficeColor.Black), OfficeImageExportFormat.Png);
        Assert.Throws<InvalidDataException>(() => PdfDocument.Load(source).Annotations.AddStamp(
            new PdfStampAnnotationOptions { ImageBytes = image, MaximumEncodedImageBytes = image.Length - 1 }));
        Assert.Empty(PdfInspector.Inspect(source).Annotations);
    }
    [Fact]
    public void AddStampAnnotation_CreatesVisualAppearanceDuringFullRewrite() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Stamp annotation source"))
            .ToBytes();

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddStampAnnotation(
            source,
            new PdfStampAnnotationOptions {
                X = 72,
                Y = 640,
                Width = 180,
                Height = 54,
                StampName = "TopSecret",
                Contents = "Reviewed stamp",
                Title = "OfficeIMO reviewer",
                Name = "review-stamp-1",
                FillColor = new PdfColor(1, 0.95, 0.9)
            });

        PdfAnnotation stamp = Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Stamp"));
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.MutationPlan.ExecutionMode);
        Assert.True(result.RewritePreservationReport!.IsPreserved);
        Assert.Equal("Reviewed stamp", stamp.Contents);
        Assert.Equal("OfficeIMO reviewer", stamp.Title);
        Assert.Equal("review-stamp-1", stamp.Name);
        Assert.Equal(72, stamp.X1);
        Assert.Equal(640, stamp.Y1);
        Assert.Equal(252, stamp.X2);
        Assert.Equal(694, stamp.Y2);
        Assert.True(stamp.HasNormalAppearance);
        Assert.Equal(new[] { 0.7D, 0.05D, 0.05D }, stamp.Color);
        Assert.Contains("/Subtype /Form", PdfEncoding.Latin1GetString(result.Bytes), StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Helvetica", PdfEncoding.Latin1GetString(result.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void CertifiedP3StampAnnotationUsesAppendOnlyRevision() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Certified stamp source"))
            .ToBytes();
        byte[] certified = Certify(source, PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures);

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddStampAnnotation(
            certified,
            new PdfStampAnnotationOptions {
                StampName = "Approved",
                Contents = "Approved after certification"
            });

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.True(result.Bytes.AsSpan(0, certified.Length).SequenceEqual(certified));
        PdfAnnotation stamp = Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Stamp"));
        Assert.Equal("Approved after certification", stamp.Contents);
        Assert.True(stamp.HasNormalAppearance);
    }

    [Fact]
    public void CertifiedP2BlocksStampAnnotationCreation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Restricted stamp source"))
            .ToBytes();
        byte[] certified = Certify(source, PdfCertificationPermissionLevel.FormFillingAndSignatures);

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfAnnotationEditor.AddStampAnnotation(certified));

        Assert.Equal(PdfMutationExecutionMode.Blocked, exception.Plan.ExecutionMode);
        Assert.Contains("AppendOnly.ActionBlocked.Annotations", exception.Plan.BlockerCodes);
    }

    [Fact]
    public void EncryptedStampAnnotationUsesAuthenticatedAppendContext() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted stamp source"))
            .ToBytes();
        var readOptions = new PdfLoadOptions { Password = "owner" };

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddStampAnnotation(
            source,
            new PdfStampAnnotationOptions {
                StampName = "Approved",
                Contents = "Encrypted approval"
            },
            readOptions);

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.True(result.Bytes.AsSpan(0, source.Length).SequenceEqual(source));
        Assert.Equal("Encrypted approval", Assert.Single(PdfInspector.Inspect(result.Bytes, readOptions).GetAnnotationsBySubtype("Stamp")).Contents);
    }

    [Fact]
    public void AddStampAnnotation_ReservesGeneratedAnnotationAndAppearanceLimits() {
        byte[] source = PdfDocument.Create()
            .TextAnnotation("Existing note")
            .Paragraph(paragraph => paragraph.Text("Tight stamp budget"))
            .ToBytes();
        int maximumSourceStreamBytes = PdfSyntax.ParseObjects(source).Map.Values
            .Select(static indirect => indirect.Value)
            .OfType<PdfStream>()
            .Max(static stream => stream.Data.Length);
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxAnnotationsPerPage = 1,
                MaxRawStreamBytes = maximumSourceStreamBytes,
                MaxDecodedStreamBytes = maximumSourceStreamBytes
            }
        };

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddStampAnnotation(
            source,
            new PdfStampAnnotationOptions { Contents = new string('A', 400) },
            readOptions);

        Assert.Equal(2, result.ToDocument().Reader.Annotations().Count);
    }

    [Fact]
    public void StampRejectsObjectNumberExhaustionBeforeGeneratingReferences() {
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "2147483644 0 obj", "<< /Unused true >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));

        Assert.Throws<OverflowException>(() => PdfAnnotationEditor.AddStampAnnotation(source));
    }

    private static byte[] Certify(byte[] source, PdfCertificationPermissionLevel permission) {
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Certification,
                CertificationPermission = permission,
                FieldName = "Certification",
                ReservedSignatureContentsBytes = 512
            });
        return PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });
    }
}
