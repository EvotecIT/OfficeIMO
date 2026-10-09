using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    public string FormOcrTitleLabel => _localizer.GetOrDefault("FormOcr.Title", "Recognize existing form values");
    public string FormOcrStartHintLabel => StudioOcrProvider.UnavailableReason ?? _localizer.GetOrDefault("FormOcr.StartHint", "Recognize visible text inside existing fields. Apply current drafts first. OCR languages follow the OCR workbench; recognition does not create fields or save the document.");
    public string FormOcrRecognizeLabel => _localizer.GetOrDefault("FormOcr.Recognize", "Recognize form values");
    public string FormOcrApplyLabel => _localizer.GetOrDefault("FormOcr.Apply", "Apply accepted values");
    public string FormOcrCancelLabel => _localizer.GetOrDefault("FormOcr.Cancel", "Cancel recognition / close review");
    private CancellationTokenSource? _formOcrCancellation;
    private PdfWorkspace? _formOcrWorkspace;
    private PdfWorkspaceFormOcrReview? _formOcrPreparation;
    [ObservableProperty] private FormOcrReviewViewModel? _formOcrReview;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanCancelOperation))]
    private bool _isFormOcrBusy;
    [ObservableProperty] private string? _formOcrError;
    public bool HasFormOcrReview => FormOcrReview is not null;
    public bool CanRecognizeFormValues => StudioOcrProvider.UnavailableReason is null && !_disposed && !IsFormOcrBusy && !IsWorkspaceBusy && !IsOpening &&
        _workspace?.CanFillForms == true && HasFormFields && !HasFormDrafts;
    public bool CanApplyFormOcr => !IsWorkspaceBusy && !IsFormOcrBusy && !HasFormDrafts && FormOcrReview?.CanApply == true &&
        _formOcrPreparation is { } preparation && ReferenceEquals(preparation.Owner, _workspace) && preparation.Revision == _workspace.Revision;

    [RelayCommand]
    private async Task RecognizeFormValuesAsync() {
        if (!CanRecognizeFormValues) return;
        await PrepareFormOcrAsync(async token => {
            var languages = OcrWorkbench.Languages.Where(choice => choice.IsSelected)
                .Aggregate((TesseractOcrLanguage)0, (current, choice) => current | choice.Value);
            if (languages == 0) throw new InvalidOperationException(_localizer.GetOrDefault("FormOcr.ChooseLanguage", "Choose an OCR language in the OCR workbench first."));
            return await _services.Ocr.CreateEngineAsync(languages, OcrWorkbench.ProvisionMissingLanguageData, token).ConfigureAwait(false);
        });
    }

    internal Task PrepareFormOcrAsync(IOcrEngine engine) => PrepareFormOcrAsync(_ => Task.FromResult(engine));

    private async Task PrepareFormOcrAsync(Func<CancellationToken, Task<IOcrEngine>> createEngine) {
        if (!CanRecognizeFormValues || _workspace is not { } workspace) return;
        ClearFormOcrReview();
        using var cancellation = new CancellationTokenSource();
        _formOcrCancellation = cancellation; _formOcrWorkspace = workspace; IsFormOcrBusy = true; FormOcrError = null;
        NotifyFormOcrState();
        try {
            using var execution = await _services.Jobs.EnterAsync(cancellation.Token);
            var engine = await createEngine(cancellation.Token);
            var preparation = await workspace.PrepareFormOcrAsync(engine, new PdfOcrMergeOptions {
                SourceName = workspace.FileName, MinimumConfidence = 0.5, MaxPages = 100
            }, cancellation.Token);
            cancellation.Token.ThrowIfCancellationRequested();
            if (_disposed || !ReferenceEquals(_workspace, workspace) || workspace.Revision != preparation.Revision) return;
            _formOcrPreparation = preparation;
            FormOcrReview = new(preparation.Review, _localizer);
            FormOcrReview.PropertyChanged += OnFormOcrReviewChanged;
        } catch (OperationCanceledException) { }
        catch (Exception exception) {
            if (!_disposed && ReferenceEquals(workspace, _workspace)) FormOcrError = exception.Message;
        } finally {
            if (ReferenceEquals(_formOcrCancellation, cancellation)) { _formOcrCancellation = null; _formOcrWorkspace = null; }
            IsFormOcrBusy = false; NotifyFormOcrState();
        }
    }

    [RelayCommand]
    private async Task ApplyFormOcrAsync(CancellationToken token) {
        if (!CanApplyFormOcr || _workspace is not { } workspace || _formOcrPreparation is not { } preparation || FormOcrReview is not { } review) return;
        var values = review.CaptureAccepted();
        await RunMutationAsync(cancellation => workspace.ApplyFormOcrAsync(preparation, values, cancellation, CreateProgress()), token);
        NotifyFormOcrState();
    }

    [RelayCommand]
    private void CancelFormOcr() { _formOcrCancellation?.Cancel(); ClearFormOcrReview(); NotifyFormOcrState(); }

    private void ClearFormOcrReview() {
        _formOcrCancellation?.Cancel();
        if (FormOcrReview is { } review) { review.PropertyChanged -= OnFormOcrReviewChanged; review.Dispose(); }
        FormOcrReview = null; _formOcrPreparation = null;
    }
    private void OnFormOcrReviewChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs args) =>
        OnPropertyChanged(nameof(CanApplyFormOcr));
    private void NotifyFormOcrState() {
        if (_formOcrWorkspace is not null && !ReferenceEquals(_formOcrWorkspace, _workspace)) _formOcrCancellation?.Cancel();
        if (_formOcrPreparation is { } foreign && !ReferenceEquals(foreign.Owner, _workspace)) ClearFormOcrReview();
        if (_formOcrPreparation is { } preparation && FormOcrReview is { } review)
            review.IsStale = !ReferenceEquals(preparation.Owner, _workspace) || preparation.Revision != _workspace?.Revision;
        OnPropertyChanged(nameof(CanRecognizeFormValues)); OnPropertyChanged(nameof(CanApplyFormOcr));
        OnPropertyChanged(nameof(HasFormOcrReview));
    }
}
