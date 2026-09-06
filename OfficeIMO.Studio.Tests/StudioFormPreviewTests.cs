using Avalonia;
using Avalonia.Media.Imaging;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioFormPreviewTests {
    [Fact]
    public async Task DocumentReplacementInvalidatesCompletedAndPendingCreationPreviews() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string first = Path.Combine(services.Paths.Root, "first.pdf");
            string second = Path.Combine(services.Paths.Root, "second.pdf");
            PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).Save(first);
            PdfDocument.Create(compose => compose.Page(page => page.Size(400, 500))).Save(second);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(first);
            await model.PreviewNewFormFieldCommand.ExecuteAsync(null);
            Assert.True(model.HasFormPreview);
            await model.OpenDocumentAsync(second);
            Assert.False(model.HasFormPreview);
            Task pending = model.PreviewNewFormFieldCommand.ExecuteAsync(null);
            await model.OpenDocumentAsync(first);
            await pending;
            Assert.False(model.HasFormPreview);
            Assert.False(model.IsFormPreviewBusy);
            Assert.Null(model.FormPreviewError);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task CreationPreviewWorksBeforeAnyFieldsExistAndTracksChangedOptions() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "empty.pdf");
            PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(source);
            model.NewFormFieldName = "Created";
            model.NewFormFieldValue = "New value";
            await model.PreviewNewFormFieldCommand.ExecuteAsync(null);
            Assert.Null(model.FormPreviewError);
            Assert.True(model.HasFormPreview);
            Assert.Empty(model.FormFields);
            Assert.False(model.IsDirty);
            model.NewFormFieldWidth = 220;
            Assert.False(model.HasFormPreview);
            await model.PreviewNewFormFieldCommand.ExecuteAsync(null);
            Assert.True(model.HasFormPreview);
            byte[] preview = Encode(model.FormPreviewImage!);
            await model.CreateFormFieldCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.False(model.HasFormPreview);
            Assert.Equal("Created", Assert.Single(model.FormFields).Name);
            await model.SaveCommand.ExecuteAsync(null);
            var rendered = PdfDocument.Load(source).Render.Pages("1", new PdfPageRenderOptions { Format = PdfPageRenderFormat.Png, Scale = 1D }).Single();
            using var stream = new MemoryStream(rendered.Bytes!);
            using var bitmap = new Bitmap(stream);
            Assert.Equal(preview, Encode(bitmap));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task PreviewMatchesSavedAppearanceAndChangedValuesInvalidatePendingResults() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "preview.pdf");
            string output = Path.Combine(services.Paths.Root, "applied.pdf");
            byte[] original = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400)))
                .Forms.Edit(edit => edit.Create(new() { Name = "Reference", Value = "Original", X = 30, Y = 300 })).ToBytes();
            File.WriteAllBytes(source, original);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(output));
            await model.OpenDocumentAsync(source);
            model.SelectedFormField!.TextValue = "Preview value";
            await model.PreviewFormAppearanceCommand.ExecuteAsync(null);
            Assert.Null(model.FormPreviewError);
            Assert.True(model.HasFormPreview);
            byte[] preview = Encode(model.FormPreviewImage!);
            Assert.False(model.CanUndo);
            Assert.Equal(original, File.ReadAllBytes(source));
            await model.SaveAsCommand.ExecuteAsync(null);
            Assert.False(model.HasFormPreview);
            var rendered = PdfDocument.Load(output).Render.Pages("1", new PdfPageRenderOptions { Format = PdfPageRenderFormat.Png, Scale = 1D }).Single();
            Assert.True(rendered.Succeeded);
            using (var stream = new MemoryStream(rendered.Bytes!))
            using (var bitmap = new Bitmap(stream)) Assert.Equal(preview, Encode(bitmap));

            model.SelectedFormField!.TextValue = "First pending";
            Task pending = model.PreviewFormAppearanceCommand.ExecuteAsync(null);
            model.SelectedFormField.TextValue = "Newer draft";
            await pending;
            Assert.False(model.HasFormPreview);
            Assert.False(model.IsFormPreviewBusy);
            Assert.Equal("Newer draft", model.SelectedFormField.TextValue);
            Assert.Equal("Preview value", PdfDocument.Load(output).Inspect().FormFieldsByName["Reference"].Value);
            await model.PreviewFormAppearanceCommand.ExecuteAsync(null);
            Assert.True(model.HasFormPreview);
            model.ClearFormPreviewCommand.Execute(null);
            Assert.False(model.HasFormPreview);
            Assert.Equal("Newer draft", model.SelectedFormField.TextValue);
            return true;
        }, CancellationToken.None);
    }

    private static byte[] Encode(Bitmap bitmap) {
        using var stream = new MemoryStream();
        bitmap.Save(stream, PngBitmapEncoderOptions.Default);
        return stream.ToArray();
    }
}
