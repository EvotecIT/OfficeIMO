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
            new PdfAnnotationUpdateOptions {
                Contents = "Updated after certification",
                AllowResidualDataInAppendOnly = true
            });

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
            new PdfAnnotationRemovalOptions {
                Subtype = "Text",
                AllowResidualDataInAppendOnly = true
            });

        Assert.True(result.Applied);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.Empty(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Text"));
    }

    [Fact]
    public void CertifiedP3AnnotationRemovalRejectsResidualDataByDefault() {
        (byte[] signedPdf, _) = BuildCertifiedAnnotatedPdf(
            PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures);

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfAnnotationEditor.RemoveAnnotations(
                signedPdf,
                new PdfAnnotationRemovalOptions { Subtype = "Text" }));

        Assert.Contains("prior revisions", exception.Message, StringComparison.OrdinalIgnoreCase);
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

    [Fact]
    public void UnsignedMarkupUpdateRewritesGeometryAndRegeneratesAppearance() {
        byte[] pdf = PdfDocument.Create()
            .HighlightAnnotation("Review", width: 120, height: 14)
            .Paragraph(paragraph => paragraph.Text("Source"))
            .ToBytes();
        int objectNumber = Assert.Single(PdfInspector.Inspect(pdf).GetAnnotationsBySubtype("Highlight")).ObjectNumber!.Value;

        PdfAnnotationEditResult result = PdfAnnotationEditor.UpdateAnnotation(pdf, objectNumber, new PdfAnnotationUpdateOptions {
            Rectangle = new[] { 20D, 30D, 140D, 50D },
            QuadPoints = new[] { 20D, 50D, 140D, 50D, 20D, 30D, 140D, 30D },
            RegenerateAppearance = true
        });
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Highlight"));

        Assert.Equal(new[] { 20D, 30D, 140D, 50D }, new[] { annotation.X1, annotation.Y1, annotation.X2, annotation.Y2 });
        Assert.Equal(new[] { 20D, 50D, 140D, 50D, 20D, 30D, 140D, 30D }, annotation.QuadPoints);
        Assert.True(annotation.HasNormalAppearance);
    }

    [Fact]
    public void CertifiedP3GeometryAndAppearanceUpdateUsesAppendOnlyRevision() {
        (byte[] signedPdf, int annotationObjectNumber) = BuildCertifiedHighlightPdf();

        PdfAnnotationEditResult result = PdfAnnotationEditor.UpdateAnnotation(signedPdf, annotationObjectNumber, new PdfAnnotationUpdateOptions {
            Rectangle = new[] { 24D, 40D, 144D, 56D },
            QuadPoints = new[] { 24D, 56D, 144D, 56D, 24D, 40D, 144D, 40D },
            RegenerateAppearance = true,
            AllowResidualDataInAppendOnly = true
        });

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.True(HasExactPrefix(result.Bytes, signedPdf));
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Highlight"));
        Assert.True(annotation.HasNormalAppearance);
        Assert.Equal(24D, annotation.X1);
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

    private static (byte[] Pdf, int AnnotationObjectNumber) BuildCertifiedHighlightPdf() {
        byte[] source = PdfDocument.Create().HighlightAnnotation("Review", 120, 14).Paragraph(p => p.Text("Source")).ToBytes();
        int objectNumber = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Highlight")).ObjectNumber!.Value;
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(source, new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            CertificationPermission = PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures,
            FieldName = "Certification",
            ReservedSignatureContentsBytes = 512
        });
        return (PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 }), objectNumber);
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
