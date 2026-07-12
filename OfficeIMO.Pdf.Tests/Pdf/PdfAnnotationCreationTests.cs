using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAnnotationCreationTests {
    [Fact]
    public void AddAnnotation_CreatesLineGeometryAppearanceAndPopupOnExistingPage() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Existing page")).ToBytes();

        PdfAnnotationEditResult result = PdfDocument.Open(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Line",
            Rectangle = new[] { 40D, 50D, 180D, 100D },
            Line = new[] { 40D, 50D, 180D, 100D },
            LineStartEnding = "OpenArrow",
            LineEndEnding = "ClosedArrow",
            Contents = "Review line",
            Color = new[] { 0.8D, 0.1D, 0.1D },
            CreatePopup = true,
            PopupOpen = true
        });
        PdfDocumentInfo info = PdfInspector.Inspect(result.Bytes);
        PdfAnnotation line = Assert.Single(info.GetAnnotationsBySubtype("Line"));

        Assert.Equal(new[] { 40D, 50D, 180D, 100D }, line.LineCoordinates);
        Assert.Equal("OpenArrow", line.LineStartEnding);
        Assert.Equal("ClosedArrow", line.LineEndEnding);
        Assert.True(line.HasNormalAppearance);
        Assert.Single(info.GetAnnotationsBySubtype("Popup"));
        Assert.Contains("/Open true", Encoding.ASCII.GetString(result.Bytes), StringComparison.Ordinal);
        Assert.NotNull(result.RewritePreservationReport);
    }

    [Fact]
    public void AddAnnotation_CreatesReplyRelationship() {
        byte[] source = PdfDocument.Create().TextAnnotation("Parent").Paragraph(p => p.Text("Existing page")).ToBytes();
        int parentObject = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddAnnotation(source, new PdfAnnotationCreateOptions {
            Subtype = "Text",
            Rectangle = new[] { 70D, 70D, 90D, 90D },
            Contents = "Reply",
            InReplyToObjectNumber = parentObject,
            ReplyType = "R",
            IconName = "Comment"
        });

        string raw = Encoding.ASCII.GetString(result.Bytes);
        Assert.Equal(2, PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Text").Count);
        Assert.Contains("/IRT ", raw, StringComparison.Ordinal);
        Assert.Contains("/RT /R", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void AddAnnotation_UsesAppendOnlyRevisionWhenCertificationAllowsAnnotations() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Certified page")).ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(source, new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            CertificationPermission = PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures,
            FieldName = "Certification",
            ReservedSignatureContentsBytes = 512
        });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddAnnotation(signed, new PdfAnnotationCreateOptions {
            Subtype = "Text",
            Contents = "Append-only review",
            IconName = "Note"
        });

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.True(result.Bytes.AsSpan(0, signed.Length).SequenceEqual(signed));
        Assert.Equal("Append-only review", Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Text")).Contents);
    }

    [Fact]
    public void FlattenAnnotations_FlattensOnlySelectedObjectThroughFluentSurface() {
        byte[] source = PdfDocument.Create()
            .FreeTextAnnotation("Flatten me", 120, 30)
            .HighlightAnnotation("Keep me", 120, 14)
            .Paragraph(p => p.Text("Existing page"))
            .ToBytes();
        int freeTextObject = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("FreeText")).ObjectNumber!.Value;

        PdfAnnotationEditResult result = PdfDocument.Open(source).Annotations.Flatten(new PdfAnnotationFlattenOptions { ObjectNumber = freeTextObject });
        PdfDocumentInfo info = PdfInspector.Inspect(result.Bytes);

        Assert.Equal(1, result.AffectedAnnotationCount);
        Assert.Empty(info.GetAnnotationsBySubtype("FreeText"));
        Assert.Single(info.GetAnnotationsBySubtype("Highlight"));
        Assert.NotNull(result.RewritePreservationReport);
    }
}
