using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAnnotationReviewWorkflowTests {
    [Fact]
    public void AddReplyAndSetState_RoundTripTypedThreadMetadata() {
        byte[] source = PdfDocument.Create()
            .TextAnnotation("Parent note")
            .Paragraph(paragraph => paragraph.Text("Review page"))
            .ToBytes();
        int parentObjectNumber = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;

        PdfAnnotationEditResult replyResult = PdfAnnotationReviewEditor.AddReply(
            source,
            parentObjectNumber,
            "Reply text",
            new PdfAnnotationReplyOptions {
                Author = "Reviewer",
                Subject = "Evidence",
                ReviewState = PdfAnnotationReviewState.None
            });
        PdfAnnotation reply = Assert.Single(
            PdfInspector.Inspect(replyResult.Bytes).Annotations,
            annotation => annotation.Review?.InReplyToObjectNumber == parentObjectNumber);
        PdfAnnotationEditResult stateResult = PdfAnnotationReviewEditor.SetState(
            replyResult.Bytes,
            reply.ObjectNumber!.Value,
            PdfAnnotationReviewState.Accepted);
        PdfAnnotationReviewCatalog catalog = PdfAnnotationReviewCatalog.Read(stateResult.Bytes);

        PdfAnnotationReviewThread thread = Assert.Single(catalog.Threads, item => item.Root.Annotation.ObjectNumber == parentObjectNumber);
        PdfAnnotationReviewEntry replyEntry = Assert.Single(thread.Root.Replies);
        Assert.Equal("Reply text", replyEntry.Annotation.Contents);
        Assert.Equal("Reviewer", replyEntry.Annotation.Title);
        Assert.Equal("Evidence", replyEntry.Annotation.Review!.Subject);
        Assert.Equal("R", replyEntry.Annotation.Review.ReplyType);
        Assert.Equal(PdfAnnotationReviewState.Accepted, replyEntry.Annotation.Review.StandardState);
        Assert.Equal(2, catalog.AnnotationCount);
        Assert.Equal(1, catalog.ReplyCount);
        Assert.Equal(0, catalog.OrphanedReplyCount);
    }

    [Fact]
    public void SetState_UsesAppendOnlyMutationWhenCertificationRequiresIt() {
        byte[] source = PdfDocument.Create().TextAnnotation("Certified note").Paragraph(paragraph => paragraph.Text("Certified page")).ToBytes();
        int objectNumber = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(source, new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            CertificationPermission = PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures,
            FieldName = "ReviewCertification",
            ReservedSignatureContentsBytes = 512
        });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });

        PdfAnnotationEditResult result = PdfAnnotationReviewEditor.SetState(
            signed,
            objectNumber,
            PdfAnnotationReviewState.Completed,
            allowResidualDataInAppendOnly: true);

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.True(result.Bytes.AsSpan(0, signed.Length).SequenceEqual(signed));
        Assert.Equal(PdfAnnotationReviewState.Completed, Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Text")).Review!.StandardState);
    }

    [Fact]
    public void FullRewrite_TracksSparseAnnotationObjectNumbersAcrossReplyAndStateReadback() {
        byte[] source = BuildSparseAnnotationPdf();

        PdfAnnotationEditResult state = PdfAnnotationReviewEditor.SetState(source, 20, PdfAnnotationReviewState.Accepted);
        PdfAnnotation parent = Assert.Single(PdfInspector.Inspect(state.Bytes).GetAnnotationsBySubtype("Text"));
        Assert.NotEqual(20, parent.ObjectNumber);
        Assert.Equal(PdfAnnotationReviewState.Accepted, parent.Review!.StandardState);

        PdfAnnotationEditResult reply = PdfAnnotationReviewEditor.AddReply(source, 20, "Sparse reply");
        PdfAnnotationReviewCatalog catalog = PdfAnnotationReviewCatalog.Read(reply.Bytes);
        PdfAnnotationReviewThread thread = Assert.Single(catalog.Threads, item => item.Root.Annotation.Contents == "Sparse parent");
        Assert.Equal("Sparse reply", Assert.Single(thread.Root.Replies).Annotation.Contents);
    }

    [Fact]
    public void Build_FailsClosedBeforeDeepReplyChainsCanExhaustTheStack() {
        var annotations = new List<PdfAnnotation>();
        for (int objectNumber = 1; objectNumber <= 130; objectNumber++) {
            PdfAnnotationReviewInfo? review = objectNumber == 1
                ? null
                : new PdfAnnotationReviewInfo(objectNumber - 1, "R", null, null, null, null);
            annotations.Add(new PdfAnnotation(
                objectNumber,
                pageNumber: 1,
                subtype: "Text",
                contents: objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                x1: 0,
                y1: 0,
                x2: 18,
                y2: 18,
                hasNormalAppearance: false,
                review: review));
        }

        Assert.Throws<InvalidOperationException>(() => PdfAnnotationReviewCatalog.Build(annotations));
        PdfAnnotationReviewCatalog catalog = PdfAnnotationReviewCatalog.Build(
            annotations,
            new PdfAnnotationReviewCatalogOptions { MaximumThreadDepth = 256 });
        Assert.Equal(130, catalog.AnnotationCount);
    }

    private static byte[] BuildSparseAnnotationPdf() => Encoding.ASCII.GetBytes(
        "%PDF-1.7\n" +
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
        "2 0 obj\n<< /Type /Pages /Count 1 /Kids [10 0 R] >>\nendobj\n" +
        "10 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 30 0 R /Annots [20 0 R] >>\nendobj\n" +
        "20 0 obj\n<< /Type /Annot /Subtype /Text /Rect [36 36 54 54] /Contents (Sparse parent) /P 10 0 R >>\nendobj\n" +
        "30 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
        "trailer\n<< /Root 1 0 R /Size 31 >>\nstartxref\n0\n%%EOF\n");
}
