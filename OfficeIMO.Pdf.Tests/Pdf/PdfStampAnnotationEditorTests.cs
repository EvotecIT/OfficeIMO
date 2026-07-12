using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfStampAnnotationEditorTests {
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
        var readOptions = new PdfReadOptions { Password = "owner" };

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
