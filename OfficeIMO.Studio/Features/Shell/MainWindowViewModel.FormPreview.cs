using Avalonia.Media.Imaging;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private CancellationTokenSource? _formPreviewCancellation;
    private long _formPreviewGeneration;
    [ObservableProperty] private Bitmap? _formPreviewImage;
    [ObservableProperty] private string? _formPreviewError;
    [ObservableProperty] private bool _isFormPreviewBusy;
    [ObservableProperty] private string _formPreviewDescription = string.Empty;

    public bool HasFormPreview => FormPreviewImage is not null;
    public bool CanPreviewFormAppearance => !IsWorkspaceBusy && !IsFormPreviewBusy &&
        _workspace?.CanFillForms == true && SelectedFormField?.CanApplyValue == true && SelectedFormField.PageNumbers.Count > 0;

    public bool CanPreviewFormPlacement => CanMoveFormWidget && !IsFormPreviewBusy;
    public bool CanPreviewNewFormField => CanAuthorForms && !IsWorkspaceBusy && !IsFormPreviewBusy;

    protected override void OnPropertyChanged(PropertyChangedEventArgs args) {
        base.OnPropertyChanged(args);
        if (args.PropertyName?.StartsWith("NewFormField", StringComparison.Ordinal) == true ||
            args.PropertyName?.StartsWith("FormWidget", StringComparison.Ordinal) == true ||
            args.PropertyName == nameof(SelectedFormFieldCreationChoice) || args.PropertyName == nameof(IsWorkspaceBusy) ||
            args.PropertyName == nameof(IsOpening)) ClearFormPreview();
    }

    partial void OnIsFormPreviewBusyChanged(bool value) => NotifyFormPreviewActions();
    private void NotifyFormPreviewActions() {
        OnPropertyChanged(nameof(CanPreviewFormAppearance));
        OnPropertyChanged(nameof(CanPreviewFormPlacement));
        OnPropertyChanged(nameof(CanPreviewNewFormField));
    }
    partial void OnFormPreviewImageChanged(Bitmap? value) => OnPropertyChanged(nameof(HasFormPreview));

    [RelayCommand]
    private void ClearFormPreview() {
        ++_formPreviewGeneration;
        _formPreviewCancellation?.Cancel();
        _formPreviewCancellation = null;
        Bitmap? previous = FormPreviewImage;
        FormPreviewImage = null;
        previous?.Dispose();
        FormPreviewError = null;
        IsFormPreviewBusy = false;
        NotifyFormPreviewActions();
    }

    [RelayCommand]
    private async Task PreviewFormAppearanceAsync(CancellationToken cancellationToken) {
        if (!CanPreviewFormAppearance || SelectedFormField is null) return;
        int pageNumber = SelectedFormField.PageNumbers[0];
        var values = new Dictionary<string, PdfFormFieldValue>(StringComparer.Ordinal) { [SelectedFormField.Name] = SelectedFormField.CreateValue() };
        await RenderFormPreviewAsync(document => {
            var plan = document.PlanMutation(PdfMutationOperation.FillFormFields, values.Keys);
            return plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly
                ? document.Forms.AppendRevision(values) : document.Forms.Fill(values);
        }, pageNumber, UiText("Forms.PreviewHint"), cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task PreviewFormPlacementAsync(CancellationToken cancellationToken) {
        if (!CanPreviewFormPlacement || SelectedFormField is null) return;
        string name = SelectedFormField.Name;
        int page = FormWidgetPage;
        double x = FormWidgetX, y = FormWidgetY, width = FormWidgetWidth, height = FormWidgetHeight;
        await RenderFormPreviewAsync(document => document.Forms.Edit(edit => edit.Move(name, page, x, y, width, height)).ToDocument(),
            page, UiText("Forms.PlacementPreviewHint"), cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task PreviewNewFormFieldAsync(CancellationToken cancellationToken) {
        if (!CanPreviewNewFormField) return;
        var options = CaptureNewFormFieldOptions();
        await RenderFormPreviewAsync(document => document.Forms.Edit(edit => edit.Create(options)).ToDocument(),
            options.PageNumber, UiText("Forms.CreationPreviewHint"), cancellationToken).ConfigureAwait(true);
    }

    private async Task RenderFormPreviewAsync(Func<PdfDocument, PdfDocument> transform, int pageNumber, string description, CancellationToken cancellationToken) {
        if (_workspace is null) return;
        var workspace = _workspace;
        long revision = workspace.Revision;
        ClearFormPreview();
        FormPreviewDescription = description;
        long generation = _formPreviewGeneration;
        using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        _formPreviewCancellation = cancellation;
        var token = cancellation.Token;
        IsFormPreviewBusy = true;
        try {
            var document = workspace.CreateDocumentSnapshot();
            byte[] bytes = await Task.Run(() => {
                token.ThrowIfCancellationRequested();
                var candidate = transform(document);
                token.ThrowIfCancellationRequested();
                var page = candidate.Inspect().Pages[pageNumber - 1];
                var result = candidate.Render.Pages(pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    new PdfPageRenderOptions {
                        Format = PdfPageRenderFormat.Png, Scale = Math.Min(1D, 700D / Math.Max(page.Width, page.Height)),
                        MaxPages = 1, MaxPixelsPerPage = 1_000_000,
                        MaxOutputBytesPerPage = 8 * 1024 * 1024, MaxTotalOutputBytes = 8 * 1024 * 1024
                    }, cancellationToken: token).Single();
                if (!result.Succeeded || result.Bytes is null)
                    throw new InvalidOperationException(string.Join(Environment.NewLine, result.Diagnostics));
                return result.Bytes;
            }, token).ConfigureAwait(true);
            if (generation != _formPreviewGeneration || token.IsCancellationRequested ||
                !ReferenceEquals(_workspace, workspace) || workspace.Revision != revision) return;
            using var stream = new MemoryStream(bytes, writable: false);
            FormPreviewImage = new Bitmap(stream);
        } catch (OperationCanceledException) when (token.IsCancellationRequested) {
        } catch (Exception exception) {
            if (generation == _formPreviewGeneration && ReferenceEquals(_workspace, workspace) && workspace.Revision == revision)
                FormPreviewError = exception.Message;
        } finally {
            if (ReferenceEquals(_formPreviewCancellation, cancellation)) _formPreviewCancellation = null;
            if (generation == _formPreviewGeneration) IsFormPreviewBusy = false;
        }
    }
}
