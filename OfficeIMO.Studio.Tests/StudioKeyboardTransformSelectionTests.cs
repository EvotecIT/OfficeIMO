using Avalonia;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioKeyboardTransformSelectionTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ConsecutiveKeyboardTransformsKeepASelectionFromTheCurrentRevision(bool annotation, bool restricted) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "keyboard-transform.pdf");
            string saved = Path.Combine(services.Paths.Root, "keyboard-transform-edited.pdf");
            var document = PdfDocument.Create(compose => compose.Page(page => page.Size(400, 400)));
            if (annotation) document = document.Annotations.Add(new PdfAnnotationCreateOptions {
                PageNumber = 1, Subtype = "Square", Name = "target", Rectangle = [50D, 50D, 90D, 90D], GenerateAppearance = true
            }).ToDocument();
            else document = document.Images.Add(new PdfPageRegion(1, 50, 50, 40, 40), Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")).Document;
            byte[] original = restricted ? document.Security.Encrypt(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.ModifyAnnotations
            }).Pdf : document.ToBytes();
            File.WriteAllBytes(source, original);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => Task.FromResult<string?>(saved), promptPdfPassword: (_, _, _) => Task.FromResult<string?>("open"), services: services);
            await model.OpenDocumentAsync(source);
            if (annotation) model.ShowAnnotateModeCommand.Execute(null); else model.ShowEditModeCommand.Execute(null);
            var region = document.Render.Interactions(1).Regions.Single(item =>
                item.Kind == (annotation ? PdfInteractionKind.Annotation : PdfInteractionKind.Image));
            var initial = new PdfEditorSelection(annotation ? PdfEditorSelectionKind.Annotation : PdfEditorSelectionKind.Image,
                1, new(region.Quad.Left, region.Quad.Top, region.Quad.Right, region.Quad.Bottom),
                ObjectNumber: region.ObjectNumber, Subtype: region.Subtype, ImagePlacement: region.ImagePlacement);
            model.Pages[0].SelectObject(initial);
            for (int step = 0; step < 2; step++) {
                var selected = Assert.IsType<PdfEditorSelection>(model.SelectedObject);
                var bounds = selected.Bounds;
                model.Pages[0].TransformObject(new(selected, new(bounds.Left + 1, bounds.Top, bounds.Right + 1, bounds.Bottom),
                    model.SelectedAnnotations, KeepSelection: true));
                await TextEditingReviewTests.WaitUntilAsync(() => !model.IsWorkspaceBusy);
                Assert.False(model.HasError, model.ErrorMessage);
                Assert.Equal(initial.Bounds.Left + step + 1, Assert.IsType<PdfEditorSelection>(model.SelectedObject).Bounds.Left, 5);
            }
            await model.SaveAsCommand.ExecuteAsync(null);
            var output = PdfDocument.Load(saved, new PdfLoadOptions { Password = restricted ? "owner" : null }).Render.Interactions(1).Regions.Single(item => item.Kind == region.Kind);
            Assert.Equal(initial.Bounds.Left + 2, output.Quad.Left, 5);
            Assert.Equal(initial.Bounds.Top, output.Quad.Top, 5);
            Assert.Equal(original, File.ReadAllBytes(source));
            return true;
        }, CancellationToken.None);
    }
}
