using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAnnotationBatchTests {
    [Fact]
    public void GroupMoveCopyAndUngroupKeepGeometryAppearancesAndReplies() {
        var source = CreateAnnotatedDocument();
        int[] numbers = source.Inspect().Annotations.Select(annotation => annotation.ObjectNumber!.Value).ToArray();
        source = source.Annotations.AddReply(numbers[0], "Keep this conversation").ToDocument();
        var before = source.Inspect().Annotations;
        numbers = before.Where(annotation => annotation.Review?.IsReply != true).Select(annotation => annotation.ObjectNumber!.Value).ToArray();
        var grouped = source.Annotations.Group(numbers).ToDocument();
        var (groupedObjects, _) = PdfSyntax.ParseObjects(grouped.ToBytes());
        var catalog = groupedObjects.Values.Select(value => value.Value).OfType<PdfDictionary>().First(dictionary => dictionary.Get<PdfName>("Type")?.Name == "Catalog");
        Assert.Equal("1.6", catalog.Get<PdfName>("Version")?.Name);
        var members = grouped.Inspect().Annotations.Where(annotation => annotation.Review?.IsReply != true).ToArray();
        Assert.Equal(2, PdfAnnotationGrouping.GetMembers(grouped.Inspect().Annotations, members[1].ObjectNumber!.Value).Count);
        Assert.Single(grouped.Inspect().Annotations, annotation => annotation.Review?.IsReply == true);

        var moved = grouped.Annotations.MoveMany(new[] { members[1].ObjectNumber!.Value }, 25, -10).ToDocument();
        var movedMembers = moved.Inspect().Annotations.Where(annotation => annotation.Review?.IsReply != true).ToArray();
        Assert.All(movedMembers, annotation => Assert.True(annotation.HasNormalAppearance));
        Assert.Equal(new[] { 70D, 45D, 200D, 85D }, movedMembers.Single(annotation => annotation.Subtype == "Line").LineCoordinates);
        Assert.Equal(new[] { 65D, 140D, 165D, 140D, 65D, 120D, 165D, 120D }, movedMembers.Single(annotation => annotation.Subtype == "Highlight").QuadPoints);

        var copied = moved.Annotations.CopyMany(new[] { movedMembers[0].ObjectNumber!.Value }, 12, 16).ToDocument();
        var copyAnnotations = copied.Inspect().Annotations;
        Assert.Equal(5, copyAnnotations.Count);
        Assert.Single(copyAnnotations, annotation => annotation.Review?.IsReply == true);
        var copies = copyAnnotations.Where(annotation => annotation.Name != "line" && annotation.Name != "highlight" && annotation.Review?.IsReply != true).ToArray();
        Assert.Equal(2, copies.Length);
        Assert.Equal(2, PdfAnnotationGrouping.GetMembers(copyAnnotations, copies[0].ObjectNumber!.Value).Count);
        Assert.All(copies, annotation => Assert.True(annotation.HasNormalAppearance));
        Assert.Equal(2, copies.Select(annotation => annotation.Name).Distinct().Count());
        Assert.Equal(new[] { 82D, 61D, 212D, 101D }, copies.Single(annotation => annotation.Subtype == "Line").LineCoordinates);

        var ungrouped = copied.Annotations.Ungroup(new[] { copies[1].ObjectNumber!.Value }).ToDocument();
        Assert.Single(ungrouped.Inspect().Annotations, annotation => annotation.Review?.IsGroup == true);
        Assert.Single(ungrouped.Inspect().Annotations, annotation => annotation.Review?.IsReply == true);
    }

    [Fact]
    public void BatchDeleteRetainsCommentContentsAndUngroupsRepliesToRemovedParents() {
        var source = CreateAnnotatedDocument();
        int parent = source.Inspect().Annotations[0].ObjectNumber!.Value;
        source = source.Annotations.AddReply(parent, "Retain this comment").ToDocument();
        var numbers = source.Inspect().Annotations.Where(annotation => annotation.Review?.IsReply != true).Select(annotation => annotation.ObjectNumber!.Value).ToArray();
        var grouped = source.Annotations.Group(numbers).ToDocument();
        int member = grouped.Inspect().Annotations.First(annotation => annotation.Review?.IsGroup == true).ObjectNumber!.Value;

        var removed = grouped.Annotations.RemoveMany(new[] { member }).ToDocument();

        var comment = Assert.Single(removed.Inspect().Annotations);
        Assert.Equal("Retain this comment", comment.Contents);
        Assert.Null(comment.Review?.InReplyToObjectNumber);
    }

    [Theory]
    [InlineData(PdfAnnotationOrderChange.Raise, "highlight,line,other")]
    [InlineData(PdfAnnotationOrderChange.Lower, "line,highlight,other")]
    [InlineData(PdfAnnotationOrderChange.BringToFront, "highlight,other,line")]
    [InlineData(PdfAnnotationOrderChange.SendToBack, "line,highlight,other")]
    public void ArrangeChangesPaintingOrderWithoutChangingAnnotationContents(PdfAnnotationOrderChange change, string expected) {
        var source = CreateAnnotatedDocument().Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Square", Name = "other", Rectangle = new[] { 50D, 50D, 180D, 180D }
        }).ToDocument();
        int line = source.Inspect().Annotations.First(annotation => annotation.Name == "line").ObjectNumber!.Value;

        var result = source.Annotations.Arrange(new[] { line }, change).ToDocument();

        Assert.Equal(expected, string.Join(",", result.Inspect().Annotations.Select(annotation => annotation.Name)));
        Assert.Equal("Source content", result.Read().Text.Trim());
    }

    [Fact]
    public void GroupRejectsConversationalRepliesWithoutChangingInput() {
        var source = CreateAnnotatedDocument();
        source = source.Annotations.AddReply(source.Inspect().Annotations[0].ObjectNumber!.Value, "A reply").ToDocument();
        var before = source.ToBytes();
        var numbers = source.Inspect().Annotations.Select(annotation => annotation.ObjectNumber!.Value).ToArray();

        Assert.Throws<NotSupportedException>(() => source.Annotations.Group(numbers));
        Assert.Equal(before, source.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TaggedRemovalRequiresResidualPermissionAndPreservesSourceAndReplies(bool batch) {
        string[] values = {
            "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 7 0 R /MarkInfo << /Marked true >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Contents 4 0 R /Annots [5 0 R 6 0 R] >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Type /Annot /Subtype /Square /Rect [20 20 60 60] /Contents (removed-content) /StructParent 0 >>",
            "<< /Type /Annot /Subtype /Text /Rect [80 80 100 100] /Contents (retained-comment) /IRT 5 0 R /RT /R >>",
            "<< /Type /StructTreeRoot /K 8 0 R /ParentTree 9 0 R >>",
            "<< /Type /StructElem /S /Annot /P 7 0 R /Pg 3 0 R /K << /Type /OBJR /Obj 5 0 R /Pg 3 0 R >> >>",
            "<< /Nums [0 8 0 R] >>", "<< /Title (Tagged annotation fixture) >>"
        };
        byte[] source = PdfPageExtractor.Assemble(values.Select((value, index) => PdfPageExtractor.WrapObject(index + 1,
            System.Text.Encoding.ASCII.GetBytes(value))).ToList(), 1, 10, PdfFileVersion.Pdf17);
        var document = PdfDocument.Load(source);
        Assert.Throws<NotSupportedException>(() => batch ? document.Annotations.RemoveMany(new[] { 5 }) :
            document.Annotations.Remove(new PdfAnnotationRemovalOptions { ObjectNumber = 5 }));
        var result = batch ? document.Annotations.RemoveMany(new[] { 5 }, allowResidualDataInAppendOnly: true) :
            document.Annotations.Remove(new PdfAnnotationRemovalOptions { ObjectNumber = 5, AllowResidualDataInAppendOnly = true });
        var retained = Assert.Single(result.ToDocument().Inspect().Annotations);
        Assert.Equal("retained-comment", retained.Contents);
        Assert.Null(retained.Review?.InReplyToObjectNumber);
        Assert.Equal(source, result.Bytes.Take(source.Length));
        var (objects, _) = PdfSyntax.ParseObjects(result.Bytes);
        Assert.Contains(objects.Values.Select(value => value.Value).OfType<PdfDictionary>(), dictionary => dictionary.Get<PdfName>("Type")?.Name == "StructElem");
    }

    [Fact]
    public void EncryptedBatchMoveUsesOneAppendOnlyRevisionAndEnforcesRemovalResidualPolicy() {
        var source = CreateAnnotatedDocument();
        byte[] encrypted = source.Security.Encrypt(new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.ModifyAnnotations
        }).Pdf;
        var ownerOptions = new PdfLoadOptions { Password = "owner" };
        var userOptions = new PdfLoadOptions { Password = "open" };
        int[] numbers = PdfInspector.Inspect(encrypted, ownerOptions).Annotations.Select(annotation => annotation.ObjectNumber!.Value).ToArray();

        var result = PdfDocument.Load(encrypted, userOptions).Annotations.MoveMany(numbers, 15, 20);

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.Equal(encrypted, result.Bytes.Take(encrypted.Length));
        Assert.Equal(PdfInspector.Probe(encrypted, ownerOptions).Security.RevisionCount + 1, PdfInspector.Probe(result.Bytes, ownerOptions).Security.RevisionCount);
        Assert.Throws<NotSupportedException>(() => PdfDocument.Load(encrypted, userOptions).Annotations.RemoveMany(numbers));
    }

    [Theory]
    [InlineData(PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures, 7, true)]
    [InlineData(PdfCertificationPermissionLevel.FormFillingAndSignatures, 7, false)]
    [InlineData(PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures, 4, false)]
    public void CertifiedBatchGroupRespectsCertificationAndPreservesSignedBytes(PdfCertificationPermissionLevel permission, int versionMinor, bool allowed) {
        byte[] original = CreateAnnotatedDocument().ToBytes();
        original[7] = (byte)('0' + versionMinor);
        var preparation = PdfIncrementalUpdater.PrepareExternalSignature(original, new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification, CertificationPermission = permission,
            FieldName = "Certification", ReservedSignatureContentsBytes = 512
        });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });
        int[] numbers = PdfInspector.Inspect(signed).Annotations.Where(annotation => annotation.Subtype != "Widget")
            .Select(annotation => annotation.ObjectNumber!.Value).ToArray();
        if (!allowed) {
            if (versionMinor < 6) Assert.Throws<NotSupportedException>(() => PdfDocument.Load(signed).Annotations.Group(numbers));
            else Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.Load(signed).Annotations.Group(numbers));
            return;
        }
        var result = PdfDocument.Load(signed).Annotations.Group(numbers);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.Equal(signed, result.Bytes.Take(signed.Length));
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.Single(result.ToDocument().Inspect().Annotations, annotation => annotation.Review?.IsGroup == true);
    }

    internal static PdfDocument CreateAnnotatedDocument() {
        var document = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Source content"));
        document = document.Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Line", Name = "line", Rectangle = new[] { 40D, 50D, 180D, 100D },
            Line = new[] { 45D, 55D, 175D, 95D }, GenerateAppearance = true
        }).ToDocument();
        return document.Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Highlight", Name = "highlight", Rectangle = new[] { 40D, 130D, 140D, 150D },
            QuadPoints = new[] { 40D, 150D, 140D, 150D, 40D, 130D, 140D, 130D }, GenerateAppearance = true
        }).ToDocument();
    }
}
