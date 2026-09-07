using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDocumentViewingTests {
    [Theory]
    [InlineData(PdfStandardPermissions.None)]
    [InlineData(PdfStandardPermissions.Accessibility)]
    [InlineData(PdfStandardPermissions.Print)]
    public void RestrictedViewingPreservesRasterAndDoesNotGrantExtraction(PdfStandardPermissions permissions) {
        byte[] bytes = CreatePdf(permissions);
        var restricted = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "reader" });
        PdfDocumentViewInfo view = restricted.InspectForViewing();
        Assert.Equal(PdfPasswordAuthenticationRole.User, view.Security.PasswordAuthenticationRole);
        Assert.False(view.CanExtractContent); Assert.Null(view.LogicalContent);
        Assert.Equal(permissions.HasFlag(PdfStandardPermissions.Accessibility), view.CanExtractText);
        Assert.Single(view.Pages); Assert.Empty(view.Pages[0].Annotations); Assert.Empty(view.Pages[0].FormWidgets);
        Assert.Empty(view.Pages[0].LinkAnnotations); Assert.Empty(view.Pages[0].PageActions);
        var owner = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "owner" });
        PdfDocumentViewInfo authorized = owner.InspectForViewing();
        Assert.True(authorized.CanExtractContent); Assert.NotNull(authorized.LogicalContent);
        Assert.Equal("Private metadata", authorized.LogicalContent.Metadata.Title);
        Assert.Equal(authorized.Pages[0].Width, view.Pages[0].Width);
        PdfPageRenderResult raster = restricted.Render.DisplayPage(1);
        Assert.True(raster.Succeeded); Assert.Equal(PdfPageRenderFormat.Png, raster.Format);
        Assert.Equal(owner.Render.DisplayPage(1).Bytes, raster.Bytes);
        string? capture = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VIEWING_CAPTURE_DIR");
        if (!string.IsNullOrEmpty(capture)) {
            Directory.CreateDirectory(capture); File.WriteAllBytes(Path.Combine(capture, "view-" + permissions + ".png"), raster.Bytes!);
        }
        Assert.Equal(owner.Render.Pages("1", new PdfPageRenderOptions { ContinueOnError = false })[0].Bytes, raster.Bytes);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Inspect());
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Read());
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Render.Drawing(1));
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Render.Pages("1", new PdfPageRenderOptions { ContinueOnError = false }));
        Assert.False(restricted.PlanMutation(PdfMutationOperation.ExtractPages).CanExecute);
        Assert.False(restricted.PlanMutation(PdfMutationOperation.ChangeEncryption).CanExecute);
        if (!permissions.HasFlag(PdfStandardPermissions.Print))
            Assert.Throws<PdfPermissionDeniedException>(() => restricted.GetPrintablePageLayouts());
        // A display call must not relax cached read options for any subsequent operation.
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Render.Interactions(1));
    }

    [Theory]
    [InlineData("display")]
    [InlineData("ranges")]
    [InlineData("selection")]
    [InlineData("images")]
    public void CancellationDuringGeneratedSerializationStopsFurtherComposition(string route) {
        using var cancellation = new CancellationTokenSource();
        int callbacks = 0;
        var options = new PdfOptions {
            TextLineBreakCallback = text => {
                callbacks++; cancellation.Cancel(); return new[] { text.Length / 2 };
            }
        };
        var document = PdfDocument.Create(options);
        for (int index = 0; index < 30; index++) document.Paragraph(paragraph => paragraph.Text(new string('W', 600)));
        Assert.ThrowsAny<OperationCanceledException>(() => {
            switch (route) {
                case "display": document.Render.DisplayPage(1, cancellationToken: cancellation.Token); break;
                case "ranges": document.Render.Pages("1", cancellationToken: cancellation.Token); break;
                case "selection": document.Render.Pages(PdfPageSelection.Parse("1"), cancellationToken: cancellation.Token); break;
                case "images": document.Render.ExportImages(OfficeIMO.Drawing.OfficeImageExportFormat.Png, cancellationToken: cancellation.Token); break;
            }
        });
        Assert.InRange(callbacks, 1, 2);
    }

    [Fact]
    public void RestrictedViewOmitsExistingLogicalMetadataLinksAndAttachments() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        byte[] encrypted = PdfDocument.Load(source).Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None
        }).Pdf;
        var owner = PdfDocument.Load(encrypted, new PdfLoadOptions { Password = "owner" }).InspectForViewing();
        Assert.NotNull(owner.LogicalContent); Assert.NotEmpty(owner.LogicalContent.Attachments);
        Assert.NotEmpty(owner.LogicalContent.NamedDestinations); Assert.NotEmpty(owner.Pages[0].LinkAnnotations);
        var user = PdfDocument.Load(encrypted, new PdfLoadOptions { Password = "reader" }).InspectForViewing();
        Assert.Null(user.LogicalContent); Assert.Equal(owner.PageCount, user.PageCount);
        Assert.All(user.Pages, page => { Assert.Empty(page.LinkAnnotations); Assert.Empty(page.Annotations); Assert.Empty(page.FormWidgets); });
    }

    [Fact]
    public void ViewingStillRequiresAuthenticationAndHonorsCancellationAndBudgets() {
        byte[] bytes = CreatePdf(PdfStandardPermissions.None);
        Assert.Throws<PdfPasswordRequiredException>(() => PdfDocument.Load(bytes).InspectForViewing());
        Assert.Throws<PdfInvalidPasswordException>(() => PdfDocument.Load(bytes, new PdfLoadOptions { Password = "wrong" }).Render.DisplayPage(1));
        var document = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "reader" });
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => document.InspectForViewing(cancellationToken: cancellation.Token));
        Assert.ThrowsAny<OperationCanceledException>(() => document.Render.DisplayPage(1, cancellationToken: cancellation.Token));
        Assert.Throws<PdfReadLimitException>(() => document.Render.DisplayPage(1, new PdfPageDisplayOptions { MaximumPixels = 4 }));
        Assert.Throws<PdfReadLimitException>(() => document.Render.DisplayPage(1, new PdfPageDisplayOptions { MaximumOutputBytes = 8 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Render.DisplayPage(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Render.DisplayPage(2));
        var thumbnail = document.Render.DisplayPage(1, new PdfPageDisplayOptions { MaximumDimension = 80 });
        Assert.InRange(Math.Max(thumbnail.Width, thumbnail.Height), 1, 80);
    }

    private static byte[] CreatePdf(PdfStandardPermissions permissions) =>
        PdfDocument.Create(new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = permissions
        })).Meta(title: "Private metadata")
          .Paragraph(paragraph => paragraph.Text("Visible content with restricted extraction"))
          .Canvas(canvas => canvas.Image(PdfPngTestImages.CreateRgbPng(30, 90, 180), 20D, 20D, 40D, 40D)).ToBytes();
}
