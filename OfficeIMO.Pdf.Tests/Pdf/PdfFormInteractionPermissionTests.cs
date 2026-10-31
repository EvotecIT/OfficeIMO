using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFormInteractionPermissionTests {
    [Theory]
    [InlineData(PdfStandardPermissions.None, false)]
    [InlineData(PdfStandardPermissions.FillForms, true)]
    [InlineData(PdfStandardPermissions.ModifyContents | PdfStandardPermissions.ModifyAnnotations, true)]
    [InlineData(PdfStandardPermissions.CopyContents, true)]
    public void Form_metadata_and_widget_geometry_follow_form_permissions_without_enabling_content_extraction(
        PdfStandardPermissions permissions, bool canRead) {
        var original = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas =>
            canvas.TextField("Name", "Ada", global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(24))))).Pages.Rotate(90, 1);
        byte[] encrypted = original.Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = permissions
        }).Pdf;
        var document = PdfDocument.Load(encrypted, new PdfLoadOptions { Password = "reader" });
        if (!canRead) {
            Assert.Throws<PdfPermissionDeniedException>(() => document.Forms.GetForInteraction());
            Assert.Empty(document.GetEditingInteractions(1).Regions);
        } else {
            var field = Assert.Single(document.Forms.GetForInteraction().Fields);
            Assert.Equal("Name", field.Name);
            Assert.Equal("Ada", field.Value);
            var map = document.GetEditingInteractions(1);
            var widget = Assert.Single(map.Regions);
            Assert.Equal(PdfInteractionKind.FormWidget, widget.Kind);
            Assert.Equal("Name", widget.FieldName);
            Assert.Equal(24D, widget.Quad.Width, 5);
            Assert.Equal(200D, widget.Quad.Height, 5);
            Assert.Empty(map.TextRegions);
            Assert.Null(widget.Text);
            Assert.Null(widget.Target);
        }
        if ((permissions & PdfStandardPermissions.CopyContents) == 0)
            Assert.Throws<PdfPermissionDeniedException>(() => document.Inspect());
        Assert.Equal(encrypted, document.ToBytes());
    }

    [Fact]
    public void Authorized_form_authoring_does_not_expose_restricted_preservation_inspection() {
        var original = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas =>
            canvas.TextField("Name", "Ada", global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(24)))));
        byte[] encrypted = original.Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.ModifyContents | PdfStandardPermissions.ModifyAnnotations
        }).Pdf;
        var options = new PdfLoadOptions { Password = "reader" };
        var document = PdfDocument.Load(encrypted, options);
        var edited = document.Forms.Edit(edit => edit.Rename("Name", "Recipient"));
        Assert.True(edited.PreservationReport.IsPreserved);
        Assert.False(edited.PreservationReport.CanReadDetails);
        Assert.Throws<PdfPermissionDeniedException>(() => edited.PreservationReport.Original);
        Assert.Throws<PdfPermissionDeniedException>(() => edited.PreservationReport.Rewritten);
        var reopened = PdfDocument.Load(edited.ToBytes(), options);
        Assert.Equal("Recipient", Assert.Single(reopened.Forms.GetForInteraction().Fields).Name);
        Assert.Throws<PdfPermissionDeniedException>(() => reopened.Inspect());
        Assert.Throws<PdfPermissionDeniedException>(() => PdfRewritePreservation.Assess(encrypted, edited.ToBytes(),
            new() { OriginalReadOptions = options, RewrittenReadOptions = options }));
    }
}
