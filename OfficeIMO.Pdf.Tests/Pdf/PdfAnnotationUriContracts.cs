using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAnnotationUriContracts {
    [Fact]
    public void PrimaryUriKeepsIdentityAndTypedActionsWithoutBypassingActiveContentGuard() {
        byte[] source = BuildActionPdf();
        var annotations = PdfDocument.Load(source).Reader.Annotations();
        PdfAnnotation selected = Assert.Single(annotations, item => item.Uri == "https://example.test/first");
        Assert.Equal(5, selected.ObjectNumber);
        Assert.Equal(1, selected.PageNumber);
        Assert.Equal("Link", selected.Subtype);
        Assert.Equal("URI", selected.ActionType);
        Assert.Equal("guides/second.html", Assert.Single(annotations, item => item.ObjectNumber == 6).Uri);
        Assert.Null(Assert.Single(annotations, item => item.ObjectNumber == 7).Uri); // GoTo with a stray URI key
        PdfAnnotation additional = Assert.Single(annotations, item => item.ObjectNumber == 8);
        Assert.Null(additional.Uri);
        Assert.Equal("URI", Assert.Single(additional.AdditionalActions).ActionType);
        PdfAnnotation chained = Assert.Single(annotations, item => item.ObjectNumber == 9);
        Assert.Null(chained.Uri);
        Assert.Equal("Named", chained.ActionType);
        Assert.Equal("URI", Assert.Single(chained.ChainedActions).ActionType);
        Assert.Null(Assert.Single(annotations, item => item.ObjectNumber == 10).Uri); // empty URI
        Assert.Null(Assert.Single(annotations, item => item.ObjectNumber == 11).Uri); // missing /S

        Assert.Throws<PdfMutationBlockedException>(() => PdfAnnotationEditor.UpdateAnnotation(source,
            selected.ObjectNumber!.Value, new PdfAnnotationUpdateOptions { Contents = "Selected by URI" }));
        // URI projection does not relax the existing guard for this active-action fixture.
        Assert.Equal("First", Assert.Single(PdfDocument.Load(source).Reader.Annotations(),
            item => item.Uri == selected.Uri).Contents);

    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void UriSelectionMetadataUsesAnnotationMutationPermissionWithoutGrantingExtraction(bool canAnnotate) {
        byte[] source = BuildEditableLinkPdf();
        var encryption = new PdfStandardEncryptionOptions("annotation-open") {
            OwnerPassword = "annotation-owner",
            AllowedPermissions = canAnnotate ? PdfStandardPermissions.ModifyAnnotations : PdfStandardPermissions.None
        };
        byte[] encrypted = PdfSecurityEditor.Encrypt(source, encryption).Pdf;
        var read = new PdfLoadOptions { Password = "annotation-open" };
        PdfReadDocument document = PdfReadDocument.Open(encrypted, read);
        Assert.Throws<PdfPermissionDeniedException>(() => document.Pages[0].GetAnnotations());
        if (!canAnnotate) {
            Assert.Throws<PdfMutationBlockedException>(() => PdfAnnotationEditor.GetEditingMetadata(encrypted, read));
            return;
        }

        PdfAnnotation selected = Assert.Single(PdfAnnotationEditor.GetEditingMetadata(encrypted, read),
            item => item.Uri == "https://example.test/first");
        Assert.Throws<NotSupportedException>(() => PdfAnnotationEditor.UpdateAnnotation(encrypted, selected.ObjectNumber!.Value,
            new PdfAnnotationUpdateOptions { Contents = "Authorized annotation edit" }, read));
        PdfAnnotationEditResult edited = PdfAnnotationEditor.UpdateAnnotation(encrypted, selected.ObjectNumber!.Value,
            new PdfAnnotationUpdateOptions { Contents = "Authorized annotation edit", AllowResidualDataInAppendOnly = true }, read);
        Assert.Equal("Authorized annotation edit", Assert.Single(PdfAnnotationEditor.GetEditingMetadata(edited.Bytes, read),
            item => item.Uri == selected.Uri).Contents);
        Assert.Equal("Second", Assert.Single(PdfAnnotationEditor.GetEditingMetadata(edited.Bytes, read),
            item => item.Uri == "https://example.test/second").Contents);
        Assert.Throws<PdfPermissionDeniedException>(() => PdfDocument.Load(edited.Bytes, read).Reader.Annotations());
        Assert.True(PdfInspector.Inspect(edited.Bytes, read).Security.HasEncryption);
    }

    private static byte[] BuildEditableLinkPdf() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.AddText("Annotation selection")).ToBytes();
        foreach (string target in new[] { "first", "second" }) {
            source = PdfAnnotationEditor.AddAnnotation(source, new PdfAnnotationCreateOptions {
                Subtype = "Link", Rectangle = new[] { 20D, 20D, 80D, 40D },
                Contents = target == "first" ? "First" : "Second",
                LinkUri = "https://example.test/" + target
            }).Bytes;
        }
        return source;
    }

    private static byte[] BuildActionPdf() => Encoding.ASCII.GetBytes("""
        %PDF-1.7
        1 0 obj << /Type /Catalog /Pages 2 0 R /URI << /Base (https://example.test/) >> >> endobj
        2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj
        3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Annots [5 0 R 6 0 R 7 0 R 8 0 R 9 0 R 10 0 R 11 0 R] >> endobj
        5 0 obj << /Type /Annot /Subtype /Link /Rect [20 20 80 40] /Contents (First) /A 12 0 R >> endobj
        6 0 obj << /Type /Annot /Subtype /Link /Rect [20 20 80 40] /Contents (Second) /A << /S /URI /URI (guides/second.html) >> >> endobj
        7 0 obj << /Type /Annot /Subtype /Link /Rect [20 50 80 70] /Contents (Destination) /A << /S /GoTo /D [3 0 R /Fit] /URI (https://example.test/wrong-action) >> >> endobj
        8 0 obj << /Type /Annot /Subtype /Text /Rect [20 80 80 100] /Contents (Additional) /AA << /E << /S /URI /URI (https://example.test/additional) >> >> >> endobj
        9 0 obj << /Type /Annot /Subtype /Link /Rect [20 110 80 130] /Contents (Chained) /A << /S /Named /N /NextPage /Next << /S /URI /URI (https://example.test/chained) >> >> >> endobj
        10 0 obj << /Type /Annot /Subtype /Link /Rect [20 140 80 160] /A << /S /URI /URI () >> >> endobj
        11 0 obj << /Type /Annot /Subtype /Link /Rect [20 170 80 190] /A << /URI (https://example.test/missing-type) >> >> endobj
        12 0 obj << /S /URI /URI 13 0 R >> endobj
        13 0 obj (https://example.test/first) endobj
        trailer << /Root 1 0 R >>
        %%EOF
        """);
}
