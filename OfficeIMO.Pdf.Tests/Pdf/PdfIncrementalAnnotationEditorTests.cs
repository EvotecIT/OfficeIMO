using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIncrementalAnnotationEditorTests {
    [Fact]
    public void CertifiedP3AnnotationUpdateUsesAppendOnlyRevision() {
        (byte[] signedPdf, int annotationObjectNumber) = BuildCertifiedAnnotatedPdf(
            PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures);

        PdfAnnotationEditResult result = PdfAnnotationEditor.UpdateAnnotation(
            signedPdf,
            annotationObjectNumber,
            new PdfAnnotationUpdateOptions { Contents = "Updated after certification" });

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.NotNull(result.SignatureMutationReport);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.True(HasExactPrefix(result.Bytes, signedPdf));
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Text"));
        Assert.Equal("Updated after certification", annotation.Contents);
    }

    [Fact]
    public void CertifiedP3AnnotationRemovalUsesAppendOnlyRevision() {
        (byte[] signedPdf, _) = BuildCertifiedAnnotatedPdf(
            PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures);

        PdfAnnotationEditResult result = PdfAnnotationEditor.RemoveAnnotations(
            signedPdf,
            new PdfAnnotationRemovalOptions { Subtype = "Text" });

        Assert.True(result.Applied);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.Empty(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Text"));
    }

    [Fact]
    public void CertifiedP2BlocksAnnotationMutation() {
        (byte[] signedPdf, int annotationObjectNumber) = BuildCertifiedAnnotatedPdf(
            PdfCertificationPermissionLevel.FormFillingAndSignatures);

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfAnnotationEditor.UpdateAnnotation(
                signedPdf,
                annotationObjectNumber,
                new PdfAnnotationUpdateOptions { Contents = "Forbidden update" }));
        PdfMutationPlan plan = exception.Plan;

        Assert.False(plan.CanExecute);
        Assert.Contains("AppendOnly.ActionBlocked.Annotations", plan.BlockerCodes);
        Assert.Contains("AppendOnly.ActionBlocked.Annotations", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UnsignedAnnotationUpdateReturnsFullRewritePreservationProof() {
        byte[] pdf = BuildAnnotatedPdf(out int annotationObjectNumber);

        PdfAnnotationEditResult result = PdfAnnotationEditor.UpdateAnnotation(
            pdf,
            annotationObjectNumber,
            new PdfAnnotationUpdateOptions { Title = "Updated reviewer" });

        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.MutationPlan.ExecutionMode);
        Assert.NotNull(result.RewritePreservationReport);
        Assert.Null(result.SignatureMutationReport);
        Assert.True(result.RewritePreservationReport!.IsPreserved);
    }

    private static (byte[] Pdf, int AnnotationObjectNumber) BuildCertifiedAnnotatedPdf(
        PdfCertificationPermissionLevel permission) {
        byte[] source = BuildAnnotatedPdf(out int annotationObjectNumber);
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Certification,
                CertificationPermission = permission,
                FieldName = "Certification",
                ReservedSignatureContentsBytes = 512
            });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });
        return (signed, annotationObjectNumber);
    }

    private static byte[] BuildAnnotatedPdf(out int annotationObjectNumber) {
        byte[] pdf = PdfDocument.Create()
            .TextAnnotation("Original review note", width: 24, height: 24)
            .Paragraph(paragraph => paragraph.Text("Annotated source"))
            .ToBytes();
        annotationObjectNumber = Assert.Single(PdfInspector.Inspect(pdf).GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;
        return pdf;
    }

    private static bool HasExactPrefix(byte[] value, byte[] prefix) {
        if (value.Length < prefix.Length) return false;
        for (int i = 0; i < prefix.Length; i++) {
            if (value[i] != prefix[i]) return false;
        }

        return true;
    }
}
