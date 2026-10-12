using Avalonia;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

/// <summary>Inspector changes are outlined before they are applied, and Cancel puts the fields back.</summary>
public sealed class StudioPendingChangeTests {
    [Fact]
    public async Task A_changed_crop_margin_is_outlined_on_the_selected_pages_until_it_is_applied_or_cancelled() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pending-crop-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "pages.pdf");
            PdfDocument.Create(document => {
                for (int page = 0; page < 4; page++) document.Page(content => content.Size(300, 400));
            }).Save(source);
            using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await document.OpenDocumentAsync(source);
            document.ShowPagesModeCommand.Execute(null);
            document.SetOrganizerSelection([document.OrganizerPages[1], document.OrganizerPages[2]]);
            Assert.False(document.HasPendingCrop);

            document.CropMargin = 36;
            Assert.True(document.HasPendingCrop);
            Assert.Equal(new double?[] { null, 36, 36, null }, document.OrganizerPages.Select(page => page.CropPreviewMargin));
            Assert.True(document.OrganizerPages[1].CropPreviewInset.Left > 0);

            // Selecting another page moves the outline with the selection.
            document.SetOrganizerSelection([document.OrganizerPages[3]]);
            Assert.Equal(new double?[] { null, null, null, 36 }, document.OrganizerPages.Select(page => page.CropPreviewMargin));

            document.CancelCropCommand.Execute(null);
            Assert.False(document.HasPendingCrop);
            Assert.Equal(MainWindowViewModel.DefaultCropMargin, document.CropMargin);
            Assert.All(document.OrganizerPages, page => Assert.Null(page.CropPreviewMargin));
            Assert.False(document.IsDirty);

            document.CropMargin = 24;
            await document.CropSelectedCommand.ExecuteAsync(null);
            Assert.True(document.IsDirty);
            Assert.True(document.CanUndo);
            Assert.False(document.HasPendingCrop);
            Assert.All(document.OrganizerPages, page => Assert.Null(page.CropPreviewMargin));

            // Leaving the Pages task hides an unapplied outline; returning shows it again until it is applied or cancelled.
            document.CropMargin = 30;
            document.ShowViewModeCommand.Execute(null);
            Assert.True(document.HasPendingCrop);
            Assert.All(document.OrganizerPages, page => Assert.Null(page.CropPreviewMargin));
            document.ShowPagesModeCommand.Execute(null);
            Assert.Equal(new double?[] { null, null, null, 30 }, document.OrganizerPages.Select(page => page.CropPreviewMargin));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task Discarding_form_changes_puts_every_field_back_to_its_saved_value() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((Avalonia.Application.Current as App)!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "discard-forms.pdf");
            File.WriteAllBytes(source, PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400)))
                .Forms.Edit(edit => edit.Create(new() { Name = "Reference", Value = "Original", X = 30, Y = 300 })).ToBytes());
            using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await document.OpenDocumentAsync(source);
            document.ShowFormsModeCommand.Execute(null);
            document.SelectedFormField!.TextValue = "Not wanted";
            Assert.True(document.HasFormDrafts);

            document.DiscardFormDraftsCommand.Execute(null);

            Assert.False(document.HasFormDrafts);
            Assert.False(document.IsDirty);
            Assert.Equal("Original", document.SelectedFormField!.TextValue);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task A_changed_image_size_is_outlined_around_its_centre_until_it_is_applied_or_cancelled() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pending-scale-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "image.pdf");
            PdfDocument.Create(document => document.Page(content => content.Size(300, 400))).Save(source);
            using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await document.OpenDocumentAsync(source);
            document.Pages[0].SelectObject(new PdfEditorSelection(PdfEditorSelectionKind.Image, 1, new PdfEditorVisualBounds(100D, 100D, 200D, 150D)));
            Assert.False(document.HasPendingImageScale);
            Assert.Null(document.Pages[0].PendingObjectPreview);

            document.ObjectScalePercent = 200;
            Assert.True(document.HasPendingImageScale);
            Assert.Equal(new Rect(50, 75, 200, 100), document.Pages[0].PendingObjectPreview);

            document.CancelImageScaleCommand.Execute(null);
            Assert.Equal(100, document.ObjectScalePercent);
            Assert.False(document.HasPendingImageScale);
            Assert.Null(document.Pages[0].PendingObjectPreview);

            // A new selection starts from its own size again.
            document.ObjectScalePercent = 50;
            document.Pages[0].SelectObject(new PdfEditorSelection(PdfEditorSelectionKind.Image, 1, new PdfEditorVisualBounds(10D, 10D, 30D, 30D)));
            Assert.Equal(100, document.ObjectScalePercent);
            Assert.Null(document.Pages[0].PendingObjectPreview);
        } finally { Directory.Delete(root, true); }
    }
}
