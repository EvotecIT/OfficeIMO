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
        int signedHeaderCount = PdfSyntax.CountIndirectObjectHeaders(signed, new PdfReadLimits());
        int signedObjectCount = PdfReadDocument.Open(signed).RawStructure().TotalObjectCount;
        int signedRevisionCount = PdfInspector.Probe(signed).Security.RevisionCount;
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxIndirectObjects = Math.Max(signedHeaderCount, signedObjectCount),
                MaxRevisions = signedRevisionCount
            }
        };

        PdfAnnotationEditResult result = PdfAnnotationReviewEditor.SetState(
            signed,
            objectNumber,
            PdfAnnotationReviewState.Completed,
            allowResidualDataInAppendOnly: true,
            readOptions: readOptions);

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.True(result.Bytes.AsSpan(0, signed.Length).SequenceEqual(signed));
        Assert.Single(result.ToDocument().Reader.Annotations());
        Assert.Equal(PdfAnnotationReviewState.Completed, Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Text")).Review!.StandardState);
    }

    [Fact]
    public void ReplyAndReviewStateUseAnnotationPermissionWithoutCopyPermission() {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            AllowedPermissions = PdfStandardPermissions.ModifyAnnotations
        };
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .TextAnnotation("Restricted parent")
            .Paragraph(paragraph => paragraph.Text("Restricted review page"))
            .ToBytes();
        var ownerReadOptions = new PdfLoadOptions { Password = "owner" };
        var userReadOptions = new PdfLoadOptions { Password = "open" };
        int parentObjectNumber = Assert.Single(
            PdfInspector.Inspect(source, ownerReadOptions).GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;

        PdfAnnotationEditResult replyResult = PdfAnnotationReviewEditor.AddReply(
            source,
            parentObjectNumber,
            "Restricted reply",
            readOptions: userReadOptions);
        PdfAnnotation reply = Assert.Single(
            PdfInspector.Inspect(replyResult.Bytes, ownerReadOptions).Annotations,
            annotation => annotation.Review?.InReplyToObjectNumber == parentObjectNumber);
        PdfAnnotationEditResult stateResult = PdfAnnotationReviewEditor.SetState(
            replyResult.Bytes,
            reply.ObjectNumber!.Value,
            PdfAnnotationReviewState.Accepted,
            allowResidualDataInAppendOnly: true,
            readOptions: userReadOptions);
        PdfAnnotation updatedReply = Assert.Single(
            PdfInspector.Inspect(stateResult.Bytes, ownerReadOptions).Annotations,
            annotation => annotation.ObjectNumber == reply.ObjectNumber);

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, replyResult.MutationPlan.ExecutionMode);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, stateResult.MutationPlan.ExecutionMode);
        Assert.True(stateResult.Bytes.AsSpan(0, replyResult.Bytes.Length).SequenceEqual(replyResult.Bytes));
        Assert.True(PdfInspector.Probe(stateResult.Bytes, ownerReadOptions).HasEncryption);
        Assert.Equal(PdfAnnotationReviewState.Accepted, updatedReply.Review!.StandardState);
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

    [Fact]
    public void Build_SeparatesGroupedAnnotationsFromReplyThreads() {
        PdfAnnotation parent = CreateReviewAnnotation(1, review: null);
        PdfAnnotation grouped = CreateReviewAnnotation(2, new PdfAnnotationReviewInfo(1, "Group", null, null, null, null));
        PdfAnnotation reply = CreateReviewAnnotation(3, new PdfAnnotationReviewInfo(1, null, null, null, null, null));
        PdfAnnotation unknown = CreateReviewAnnotation(4, new PdfAnnotationReviewInfo(1, "Other", null, null, null, null));

        PdfAnnotationReviewCatalog catalog = PdfAnnotationReviewCatalog.Build(new[] { parent, grouped, reply, unknown });

        PdfAnnotationReviewThread parentThread = Assert.Single(catalog.Threads, thread => thread.Root.Annotation.ObjectNumber == 1);
        Assert.Equal(3, Assert.Single(parentThread.Root.Replies).Annotation.ObjectNumber);
        PdfAnnotationReviewThread groupThread = Assert.Single(catalog.Threads, thread => thread.Root.Annotation.ObjectNumber == 2);
        Assert.False(groupThread.IsOrphanedReply);
        PdfAnnotationReviewThread unknownThread = Assert.Single(catalog.Threads, thread => thread.Root.Annotation.ObjectNumber == 4);
        Assert.False(unknownThread.IsOrphanedReply);
        Assert.True(grouped.Review!.IsGroup);
        Assert.False(grouped.Review.IsReply);
        Assert.True(reply.Review!.IsReply);
        Assert.False(unknown.Review!.IsReply);
        Assert.False(unknown.Review.IsGroup);
        Assert.Equal(1, catalog.ReplyCount);
        Assert.Equal(0, catalog.OrphanedReplyCount);
    }

    [Fact]
    public void Build_EnforcesRelationshipLimitWhileScanningAnnotations() {
        var annotations = new[] {
            CreateReviewAnnotation(1, review: null),
            CreateReviewAnnotation(2, new PdfAnnotationReviewInfo(1, "Group", null, null, null, null)),
            CreateReviewAnnotation(3, new PdfAnnotationReviewInfo(1, "R", null, null, null, null))
        };

        Assert.Throws<InvalidOperationException>(() => PdfAnnotationReviewCatalog.Build(
            annotations,
            new PdfAnnotationReviewCatalogOptions { MaximumRelationships = 1 }));
    }

    [Fact]
    public void ReviewState_RejectsNonTextAnnotationsAcrossCreateUpdateAndWorkflowSurfaces() {
        byte[] source = PdfDocument.Create()
            .HighlightAnnotation("Highlighted", 120, 14)
            .ToBytes();
        int highlightObjectNumber = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Highlight")).ObjectNumber!.Value;

        Assert.Throws<NotSupportedException>(() => PdfAnnotationReviewEditor.SetState(
            source,
            highlightObjectNumber,
            PdfAnnotationReviewState.Accepted));
        Assert.Throws<NotSupportedException>(() => PdfAnnotationEditor.UpdateAnnotation(
            source,
            highlightObjectNumber,
            new PdfAnnotationUpdateOptions { ReviewState = PdfAnnotationReviewState.Accepted }));
        Assert.Throws<NotSupportedException>(() => PdfAnnotationEditor.AddAnnotation(
            source,
            new PdfAnnotationCreateOptions {
                Subtype = "Highlight",
                Rectangle = new[] { 36D, 72D, 120D, 90D },
                ReviewState = PdfAnnotationReviewState.Accepted
        }));
    }

    [Fact]
    public void AddReply_ReservesStructuralReadLimitsForGeneratedAnnotationObjects() {
        byte[] source = PdfDocument.Create()
            .TextAnnotation("Parent note")
            .Paragraph(paragraph => paragraph.Text("Tight review budget"))
            .ToBytes();
        int parentObjectNumber = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;
        int sourceObjectCount = PdfReadDocument.Open(source).RawStructure().TotalObjectCount;
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxIndirectObjects = sourceObjectCount,
                MaxAnnotationsPerPage = 1
            }
        };

        PdfAnnotationEditResult result = PdfAnnotationReviewEditor.AddReply(
            source,
            parentObjectNumber,
            "Budgeted reply",
            readOptions: readOptions);

        Assert.Equal(2, result.ToDocument().Reader.Annotations().Count);
    }

    private static byte[] BuildSparseAnnotationPdf() => Encoding.ASCII.GetBytes(
        "%PDF-1.7\n" +
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
        "2 0 obj\n<< /Type /Pages /Count 1 /Kids [10 0 R] >>\nendobj\n" +
        "10 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 30 0 R /Annots [20 0 R] >>\nendobj\n" +
        "20 0 obj\n<< /Type /Annot /Subtype /Text /Rect [36 36 54 54] /Contents (Sparse parent) /P 10 0 R >>\nendobj\n" +
        "30 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
        "trailer\n<< /Root 1 0 R /Size 31 >>\nstartxref\n0\n%%EOF\n");

    private static PdfAnnotation CreateReviewAnnotation(int objectNumber, PdfAnnotationReviewInfo? review) => new PdfAnnotation(
        objectNumber,
        pageNumber: 1,
        subtype: "Text",
        contents: "annotation-" + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
        x1: 0,
        y1: 0,
        x2: 18,
        y2: 18,
        hasNormalAppearance: false,
        review: review);
}
