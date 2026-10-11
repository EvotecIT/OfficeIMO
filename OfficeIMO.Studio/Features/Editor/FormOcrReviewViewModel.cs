using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Editor;

public sealed record FormOcrSourceChoice(PdfFormOcrEvidence Evidence, string Label);

public sealed partial class FormOcrProposalViewModel : ObservableObject {
    private readonly PdfFormOcrReview _review;
    private readonly Action _changed;
    private readonly IStudioLocalizer _localizer;
    internal FormOcrProposalViewModel(PdfFormOcrReview review, PdfFormOcrProposal proposal,
        IStudioLocalizer localizer, Action changed) {
        _review = review; Proposal = proposal; _localizer = localizer; _changed = changed; _value = proposal.SuggestedValue;
    }
    internal PdfFormOcrProposal Proposal { get; }
    public string Name => Proposal.Field.AlternateName ?? Proposal.Field.Name!;
    public bool CanAccept => Proposal.CanAccept;
    public string Evidence => _localizer.FormatOrDefault("FormOcr.Evidence", "Page {0} · Confidence {1:P0}",
        string.Join(", ", Proposal.Evidence.Select(item => item.PageNumber).Distinct()), Proposal.Confidence);
    public string Warning => !CanAccept ? _localizer.GetOrDefault("FormOcr.Unsupported", "This field is read-only or has unsupported constraints. Fill it in a qualified form application.") :
        Proposal.IsAmbiguous ? _localizer.GetOrDefault("FormOcr.Ambiguous", "Competing field or widget assignments. Correct the value before accepting it.") :
        Proposal.HasLowConfidence ? _localizer.GetOrDefault("FormOcr.LowConfidence", "Low confidence. Check the original page and correct the value before accepting it.") : string.Empty;
    public string Validation => CanAccept ? string.Join(" ", _review.Assess(Proposal, CreateValue()).Issues.Select(issue =>
        _localizer.GetOrDefault("FormOcr.Constraint." + issue.Code, issue.Message))) : Warning;
    public bool IsValid => CanAccept && !_review.Assess(Proposal, CreateValue()).HasErrors;
    public string OriginalText => Proposal.SuggestedValue;
    public string CorrectedValueLabel => _localizer.GetOrDefault("FormOcr.CorrectedValue", "Reviewed field value");
    public string AcceptValueLabel => _localizer.GetOrDefault("FormOcr.AcceptValue", "I checked this value and accept it");
    public string Choices => string.Join(" · ", Proposal.Field.Options.Select(option => option.DisplayText == option.ExportValue
        ? option.ExportValue : option.DisplayText + " → " + option.ExportValue));
    [ObservableProperty] private string _value;
    [ObservableProperty] private bool _accepted;
    partial void OnValueChanged(string value) { Accepted = false; OnPropertyChanged(nameof(Validation)); OnPropertyChanged(nameof(IsValid)); _changed(); }
    partial void OnAcceptedChanged(bool value) { if (value && !CanAccept) Accepted = false; _changed(); }
    internal PdfFormFieldValue CreateValue() => Proposal.Field.AllowsMultipleSelection
        ? PdfFormFieldValue.FromValues(Value.Split('\n').Select(item => item.Trim()).Where(item => item.Length > 0).DefaultIfEmpty(string.Empty))
        : PdfFormFieldValue.From(Value);
}

/// <summary>Thin explicit acceptance and correction UI over source-bound form OCR proposals.</summary>
public sealed partial class FormOcrReviewViewModel : ObservableObject, IDisposable {
    public string ReviewHintLabel => _localizer.GetOrDefault("FormOcr.ReviewHint", "Compare each proposal with the captured source. Correct the value and explicitly accept it. Applying edits the current workspace; use Save a copy to keep the original file.");
    public string NoProposalsLabel => _localizer.GetOrDefault("FormOcr.NoProposals", "No text matched visible existing form widgets. Use manual field filling or adjust the OCR language and source.");
    public string StaleLabel => _localizer.GetOrDefault("FormOcr.Stale", "The document changed after recognition. Recognize it again before applying values.");
    public string ChooseProposalLabel => _localizer.GetOrDefault("FormOcr.ChooseProposal", "Choose a recognized field");
    public string CorrectedValueLabel => _localizer.GetOrDefault("FormOcr.CorrectedValue", "Reviewed field value");
    public string AcceptValueLabel => _localizer.GetOrDefault("FormOcr.AcceptValue", "I checked this value and accept it");
    public string SourcePreviewLabel => _localizer.GetOrDefault("FormOcr.SourcePreview", "Captured source page");
    public string FullPageLabel => _localizer.GetOrDefault("TextEdit.FullPage", "Show full page");
    public string DiagnosticsLabel => _localizer.GetOrDefault("FormOcr.Diagnostics", "Recognition diagnostics");
    private readonly PdfFormOcrReview _review;
    private readonly IStudioLocalizer _localizer;
    private CancellationTokenSource? _previewCancellation;
    private bool _disposed;
    internal FormOcrReviewViewModel(PdfFormOcrReview review, IStudioLocalizer localizer) {
        _review = review; _localizer = localizer;
        Proposals = review.Proposals.Select(proposal => new FormOcrProposalViewModel(review, proposal, localizer, Changed)).ToArray();
        SelectedProposal = Proposals.FirstOrDefault();
    }
    public IReadOnlyList<FormOcrProposalViewModel> Proposals { get; }
    public bool HasProposals => Proposals.Count > 0;
    public bool HasMultipleSources => Sources.Count > 1;
    public bool CanApply => !_disposed && !IsStale && Proposals.Any(item => item.Accepted) && Proposals.Where(item => item.Accepted).All(item => item.IsValid);
    public string Summary => _localizer.FormatOrDefault("FormOcr.Summary", "Accepted values: {0} of {1}. Review each value before applying.", Proposals.Count(item => item.Accepted), Proposals.Count);
    public string Diagnostics => string.Join(Environment.NewLine, _review.Ocr.Pages.SelectMany(page => page.Diagnostics));
    [ObservableProperty] private bool _isStale;
    [ObservableProperty] private FormOcrProposalViewModel? _selectedProposal;
    [ObservableProperty] private IReadOnlyList<FormOcrSourceChoice> _sources = [];
    [ObservableProperty] private FormOcrSourceChoice? _selectedSource;
    [ObservableProperty] private Bitmap? _preview;
    [ObservableProperty] private Avalonia.Rect? _previewRegion;
    [ObservableProperty] private bool _fullPage;
    [ObservableProperty] private string? _previewError;
    internal Task PreviewTask { get; private set; } = Task.CompletedTask;
    partial void OnIsStaleChanged(bool value) => Changed();
    partial void OnSelectedProposalChanged(FormOcrProposalViewModel? value) {
        Sources = value?.Proposal.Evidence.GroupBy(item => (item.PageNumber, item.WidgetBounds))
            .Select(group => {
                var item = group.First(); var bounds = item.WidgetBounds;
                return new FormOcrSourceChoice(item, _localizer.FormatOrDefault("Ocr.Review.Page", "Page {0}", item.PageNumber) +
                    $" · {bounds.Left:0.#}, {bounds.Top:0.#}, {bounds.Width:0.#} × {bounds.Height:0.#} pt");
            }).ToArray() ?? [];
        OnPropertyChanged(nameof(HasMultipleSources));
        SelectedSource = Sources.FirstOrDefault();
    }
    partial void OnSelectedSourceChanged(FormOcrSourceChoice? value) {
        UpdatePreviewRegion();
        if (!_disposed && value is not null) PreviewTask = LoadPreviewAsync(value);
    }
    partial void OnFullPageChanged(bool value) => UpdatePreviewRegion();
    private void UpdatePreviewRegion() {
        PreviewRegion = null;
        if (FullPage || SelectedSource is not { } source) return;
        var evidence = source.Evidence;
        var (width, height) = _review.GetPageSize(evidence.PageNumber);
        var bounds = evidence.WidgetBounds;
        double cropWidth = Math.Min(width, Math.Max(240, bounds.Width + 60));
        double cropHeight = Math.Min(height, Math.Max(100, bounds.Height + 70));
        double left = Math.Clamp(bounds.Left - 30, 0, Math.Max(0, width - cropWidth));
        double top = Math.Clamp(bounds.Top - 35, 0, Math.Max(0, height - cropHeight));
        PreviewRegion = new(left / width, top / height, cropWidth / width, cropHeight / height);
    }
    private void Changed() { OnPropertyChanged(nameof(CanApply)); OnPropertyChanged(nameof(Summary)); }
    internal IReadOnlyDictionary<PdfFormOcrProposal, PdfFormFieldValue> CaptureAccepted() {
        if (!CanApply) throw new InvalidOperationException("Review and accept valid values for the current document.");
        return Proposals.Where(item => item.Accepted).ToDictionary(item => item.Proposal, item => item.CreateValue());
    }
    private async Task LoadPreviewAsync(FormOcrSourceChoice source) {
        _previewCancellation?.Cancel();
        using var cancellation = new CancellationTokenSource();
        _previewCancellation = cancellation;
        Preview?.Dispose(); Preview = null;
        PreviewError = null;
        try {
            var page = await Task.Run(() => _review.RenderPage(source.Evidence.PageNumber, cancellation.Token), cancellation.Token);
            cancellation.Token.ThrowIfCancellationRequested();
            using var stream = new MemoryStream(page.Bytes!);
            var image = new Bitmap(stream);
            if (_disposed || !ReferenceEquals(SelectedSource, source)) { image.Dispose(); return; }
            Preview?.Dispose(); Preview = image;
        } catch (OperationCanceledException) { }
        catch (Exception exception) { if (!_disposed && ReferenceEquals(SelectedSource, source)) PreviewError = exception.Message; }
        finally { if (ReferenceEquals(_previewCancellation, cancellation)) _previewCancellation = null; }
    }
    public void Dispose() { _disposed = true; _previewCancellation?.Cancel(); Preview?.Dispose(); Preview = null; }
}
