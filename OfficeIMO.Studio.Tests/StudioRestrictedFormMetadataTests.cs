using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioRestrictedFormMetadataTests {
    [Fact]
    public async Task A_form_filling_permission_exposes_field_controls_without_copying_page_contents() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "restricted-form.pdf");
            var document = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas =>
                canvas.TextField("Name", "Ada", global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(24)))));
            byte[] original = document.Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.FillForms
            }).Pdf;
            File.WriteAllBytes(source, original);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                promptPdfPassword: (_, _, _) => Task.FromResult<string?>("reader"), services: services);
            await model.OpenDocumentAsync(source);
            model.ShowFormsModeCommand.Execute(null);
            Assert.Equal("Name", Assert.Single(model.FormFields).Name);
            Assert.Equal("Ada", model.SelectedFormField!.TextValue);
            Assert.False(model.CanFillForms); // Encrypted form rewriting still requires owner authorization.
            var window = new Window { Width = 1280, Height = 900, Content = new DocumentWorkspaceView { DataContext = model } };
            try {
                window.Show(); window.UpdateLayout();
                await model.Pages[0].EnsureRenderedAsync();
                window.UpdateLayout();
                var map = model.Pages[0].Scene!.Interactions!;
                Assert.Empty(map.TextRegions);
                Assert.Equal(PdfInteractionKind.FormWidget, Assert.Single(map.Regions).Kind);
                Assert.Equal(original, File.ReadAllBytes(source));
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrEmpty(output)) {
                    Directory.CreateDirectory(output);
                    using var frame = window.CaptureRenderedFrame();
                    Assert.NotNull(frame);
                    frame.Save(Path.Combine(output, "restricted-form-metadata.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
