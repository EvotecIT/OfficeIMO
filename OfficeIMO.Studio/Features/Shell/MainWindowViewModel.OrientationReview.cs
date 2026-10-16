using System.Collections.ObjectModel;
using System.Globalization;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>A review choice that becomes eligible for acceptance after its preview has been opened.</summary>
public sealed partial class PageOrientationChoice : ObservableObject {
    internal PageOrientationChoice(PdfPageOrientationProposal proposal, string label) { Proposal = proposal; Label = label; }
    internal PdfPageOrientationProposal Proposal { get; }
    public string Label { get; }
    public string Diagnostics => string.Join(Environment.NewLine, Proposal.Diagnostics);
    public bool CanAccept => HasReviewed && Proposal.CanApply;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(CanAccept))] private bool _hasReviewed;
    [ObservableProperty] private bool _isAccepted;
}

public sealed partial class MainWindowViewModel {
    private PdfWorkspaceOrientationReview? _orientationReview;
    public ObservableCollection<PageOrientationChoice> OrientationChoices { get; } = new();
    public bool HasOrientationReview => _orientationReview is not null;
    public bool CanDetectPageOrientation => CanRotateSelection && StudioOcrProvider.UnavailableReason is null &&
        (_hasOrientationProvider || OcrReadiness.IsReady);
    public bool CanApplyOrientationReview => HasOrientationReview && OrientationChoices.Any(choice => choice.IsAccepted && choice.CanAccept);
    public bool HasOrientationPreview => OrientationBefore is not null && OrientationAfter is not null;
    [ObservableProperty] private PageOrientationChoice? _selectedOrientationChoice;
    [ObservableProperty] private Bitmap? _orientationBefore;
    [ObservableProperty] private Bitmap? _orientationAfter;

    private readonly Func<CancellationToken, Task<IOcrEngine>> _createOrientationEngine;
    private readonly bool _hasOrientationProvider;

    private void OnOcrReadinessChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs args) {
        if (args.PropertyName == nameof(OcrReadinessViewModel.IsReady)) OnPropertyChanged(nameof(CanDetectPageOrientation));
    }
    private async Task<IOcrEngine> CreateOrientationEngineAsync(CancellationToken token) =>
        (await _services.Ocr.CreateSessionAsync(new TesseractOcrSessionOptions {
            CustomLanguageExpression = "osd", ProvisionMissingLanguageData = false
        }, token).ConfigureAwait(false)).Engine;

    partial void OnSelectedOrientationChoiceChanged(PageOrientationChoice? value) => ClearOrientationPreview();
    private void ClearOrientationPreview() {
        OrientationBefore?.Dispose(); OrientationAfter?.Dispose(); OrientationBefore = null; OrientationAfter = null;
        OnPropertyChanged(nameof(HasOrientationPreview));
    }
    private void ClearOrientationReview() {
        foreach (var choice in OrientationChoices) choice.PropertyChanged -= OnOrientationChoiceChanged;
        OrientationChoices.Clear(); _orientationReview = null; SelectedOrientationChoice = null; ClearOrientationPreview();
        OnPropertyChanged(nameof(HasOrientationReview)); OnPropertyChanged(nameof(CanApplyOrientationReview));
    }
    private void OnOrientationChoiceChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs args) =>
        OnPropertyChanged(nameof(CanApplyOrientationReview));

    [RelayCommand]
    private async Task DetectPageOrientationAsync(CancellationToken token) {
        if (_workspace is not { } workspace || !CanDetectPageOrientation) return;
        int[] pages = GetSelectedPages(); ClearOrientationReview();
        PdfWorkspaceOrientationReview? prepared = null;
        bool success = await RunStandaloneAsync(async cancellation => {
            IOcrEngine engine = await _createOrientationEngine(cancellation).ConfigureAwait(true);
            prepared = await workspace.PrepareOrientationReviewAsync(engine, pages, cancellation).ConfigureAwait(true);
        }, token).ConfigureAwait(true);
        if (!success || prepared is null || _disposed || !ReferenceEquals(workspace, _workspace) || workspace.Revision != prepared.Revision) return;
        _orientationReview = prepared;
        foreach (var proposal in prepared.Review.Pages) {
            string correction = proposal.ClockwiseCorrectionDegrees?.ToString(CultureInfo.CurrentCulture) ?? "—";
            string confidence = proposal.Confidence?.ToString("P0", CultureInfo.CurrentCulture) ?? "—";
            var choice = new PageOrientationChoice(proposal,
                string.Format(CultureInfo.CurrentCulture, _localizer.GetOrDefault("Orientation.Page", "Page {0}: {1}° clockwise · confidence {2}"),
                    proposal.PageNumber, correction, confidence));
            choice.PropertyChanged += OnOrientationChoiceChanged; OrientationChoices.Add(choice);
        }
        SelectedOrientationChoice = OrientationChoices.FirstOrDefault();
        OnPropertyChanged(nameof(HasOrientationReview)); OnPropertyChanged(nameof(CanApplyOrientationReview));
        ShowPagesModeCommand.Execute(null);
    }

    [RelayCommand]
    private async Task PreviewPageOrientationAsync(CancellationToken token) {
        if (_orientationReview is not { } review || SelectedOrientationChoice is not { } choice) return;
        ClearOrientationPreview();
        await RunStandaloneAsync(async cancellation => {
            var preview = await Task.Run(() => review.Review.PreviewAsync(choice.Proposal, cancellation), cancellation).ConfigureAwait(true);
            cancellation.ThrowIfCancellationRequested();
            if (_disposed || !ReferenceEquals(review, _orientationReview) || !ReferenceEquals(choice, SelectedOrientationChoice)) return;
            using var beforeStream = new MemoryStream(preview.GetSourcePng());
            using var afterStream = new MemoryStream(preview.GetPreparedPng());
            Bitmap before = new(beforeStream); Bitmap after;
            try { after = new(afterStream); } catch { before.Dispose(); throw; }
            OrientationBefore = before; OrientationAfter = after; choice.HasReviewed = true;
            OnPropertyChanged(nameof(HasOrientationPreview));
        }, token).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task ApplyPageOrientationReviewAsync(CancellationToken token) {
        if (_workspace is not { } workspace || _orientationReview is not { } review || !CanApplyOrientationReview) return;
        var accepted = OrientationChoices.Where(choice => choice.IsAccepted && choice.CanAccept).Select(choice => choice.Proposal).ToArray();
        await RunMutationAsync(cancellation => workspace.ApplyOrientationReviewAsync(review, accepted, cancellation, CreateProgress()),
            token, accepted.Select(proposal => proposal.PageNumber).ToArray()).ConfigureAwait(true);
    }

    [RelayCommand]
    private void DismissPageOrientationReview() => ClearOrientationReview();
}
