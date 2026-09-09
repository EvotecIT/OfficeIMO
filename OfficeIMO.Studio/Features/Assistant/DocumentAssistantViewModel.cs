using System.Collections.ObjectModel;
using System.Text.Json;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.AI;
using OfficeIMO.AI.IntelligenceX;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Assistant;

internal sealed record AssistantSource(byte[] Bytes, string Name, int? Page, Func<bool> IsCurrent, ReaderPdfOptions? ReaderOptions = null);
internal sealed class AssistantSourceUnavailableException(string message) : InvalidOperationException(message);

/// <summary>Read-only conversation belonging to one document tab and one connection revision.</summary>
internal sealed partial class DocumentAssistantViewModel : ObservableObject, IDisposable {
    private readonly Func<bool, AssistantSource> _capture;
    private readonly Action<int> _navigate;
    private readonly IStudioLocalizer _localizer;
    private readonly Func<OfficeAiExecutionProfile, OfficeAiIntelligenceXOptions, CancellationToken, Task<IOfficeAiExecutor>> _connect;
    private readonly List<(string Question, OfficeAiResult Result)> _history = [];
    private CancellationTokenSource? _operation;
    private AssistantSource? _source;
    private string? _snapshotHash;
    private int _generation;
    private bool _disposed;

    internal DocumentAssistantViewModel(StudioAiConnections connections, Func<bool, AssistantSource> capture,
        Action<int> navigate, IStudioLocalizer localizer,
        Func<OfficeAiExecutionProfile, OfficeAiIntelligenceXOptions, CancellationToken, Task<IOfficeAiExecutor>>? connect = null) {
        Connections = connections; _capture = capture; _navigate = navigate; _localizer = localizer;
        _connect = connect ?? (async (profile, options, token) => await IntelligenceXOfficeAiExecutor.ConnectAsync(profile, options, token));
        connections.Changed += OnConnectionChanged;
        connections.PropertyChanged += OnConnectionPropertyChanged;
        Status = Text("Ready", "Ask about the open PDF. Answers include source references for review.");
    }

    public StudioAiConnections Connections { get; }
    public ObservableCollection<AssistantMessage> Messages { get; } = [];
    [ObservableProperty] private string _question = string.Empty;
    [ObservableProperty] private string _status = string.Empty;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private bool _currentPageOnly;
    [ObservableProperty] private bool _allowRemoteProcessing;
    [ObservableProperty] private bool _showConnections = true;
    public bool CanAsk => !_disposed && !IsBusy && Connections.CanUse && (Connections.IsLocal || AllowRemoteProcessing)
        && !string.IsNullOrWhiteSpace(Question) && Question.Length <= 8000;

    partial void OnQuestionChanged(string value) => Refresh();
    partial void OnAllowRemoteProcessingChanged(bool value) { if (!value && IsBusy) Cancel(); Refresh(); }
    partial void OnCurrentPageOnlyChanged(bool value) => ResetContext();
    partial void OnIsBusyChanged(bool value) => Refresh();
    private void Refresh() { OnPropertyChanged(nameof(CanAsk)); AskCommand.NotifyCanExecuteChanged(); }

    [RelayCommand(CanExecute = nameof(CanAsk))]
    private async Task AskAsync() {
        if (!CanAsk) return;
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(3));
        _operation = cancellation; IsBusy = true;
        int generation = _generation;
        long connectionRevision = Connections.Revision;
        string question = Question.Trim();
        ShowConnections = false;
        IOfficeAiExecutor? executor = null;
        try {
            Status = Text("Reading", "Reading the current document snapshot…");
            AssistantSource source = _capture(CurrentPageOnly);
            _source = source;
            using var stream = new MemoryStream(source.Bytes, writable: false);
            OfficeAiDocument document = await OfficeAiDocument.ReadAsync(new OfficeDocumentReaderBuilder().AddPdfHandler(source.ReaderOptions).Build(),
                stream, source.Name, cancellationToken: cancellation.Token);
            if (!Current()) return;
            if (_snapshotHash != document.SnapshotHash) { _history.Clear(); Messages.Clear(); }
            _snapshotHash = document.SnapshotHash;
            string context = JsonSerializer.Serialize(_history.TakeLast(3).Select(entry => new {
                question = entry.Question, answer = entry.Result.Claims.Select(claim => claim.Text).Take(8)
            }));
            if (context.Length > 8000) context = string.Empty;
            executor = await _connect(Connections.Profile(), Connections.Options(), cancellation.Token);
            if (!Current()) return;
            Messages.Add(new AssistantMessage(question, [], true));
            Question = string.Empty;
            var progress = new Progress<OfficeAiProgress>(value => {
                if (Current()) Status = _localizer.FormatOrDefault("Assistant.Progress", "{0} · {1}/{2} batches", value.Stage, value.CompletedBatches, value.TotalBatches);
            });
            OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(document, new OfficeAiRequest {
                Instruction = question, ConversationContext = context,
                Pages = source.Page.HasValue ? [source.Page.Value] : [], AllowRemoteProcessing = AllowRemoteProcessing
            }, progress, cancellation.Token);
            if (!Current() || result.SnapshotHash != document.SnapshotHash) return;
            foreach (OfficeAiClaim claim in result.Claims) {
                var citations = claim.Citations.Select(citation => new AssistantCitation(citation,
                    () => { if (source.IsCurrent() && connectionRevision == Connections.Revision && citation.Page is int page) _navigate(page); }, _localizer)).ToArray();
                Messages.Add(new AssistantMessage(claim.Text, citations, false));
            }
            if (result.Claims.Count == 0) Messages.Add(new AssistantMessage(Text("NoAnswer", "No supported answer was returned for the selected evidence."), [], false));
            _history.Add((question, result));
            if (_history.Count > 3) _history.RemoveAt(0);
            while (Messages.Count > 60) Messages.RemoveAt(0);
            Status = _localizer.FormatOrDefault("Assistant.ResultStatus", "{0} · {1} requests · {2} omitted evidence items. Check source references; a matching quote does not prove the interpretation.", result.Status, result.RequestCount, result.OmittedEvidenceIds.Count);

            bool Current() => !_disposed && generation == _generation && connectionRevision == Connections.Revision
                && !cancellation.IsCancellationRequested && source.IsCurrent();
        } catch (OperationCanceledException) {
            if (!_disposed && generation == _generation) Status = Text("RequestCancelled", "Request cancelled. Late answers are discarded.");
        } catch (AssistantSourceUnavailableException exception) {
            // Capture failures originate in the local document boundary, not the remote provider.
            if (!_disposed && generation == _generation) Status = exception.Message;
        } catch (Exception) {
            if (!_disposed && generation == _generation) Status = Text("RequestFailed", "The document request failed. Check the connection and document support, then try again.");
        } finally {
            if (executor is IDisposable disposable) disposable.Dispose();
            if (ReferenceEquals(_operation, cancellation)) _operation = null;
            IsBusy = false;
        }
    }

    [RelayCommand] private void Cancel() { _operation?.Cancel(); }
    [RelayCommand] private void NewConversation() => ResetContext();
    internal void CheckSource() { if (_source is not null && !_source.IsCurrent()) ResetContext(); }
    internal void Deactivate() { _operation?.Cancel(); }
    private void OnConnectionChanged(object? sender, EventArgs e) { AllowRemoteProcessing = false; ResetContext(); }
    private void OnConnectionPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e) {
        if (e.PropertyName == nameof(StudioAiConnections.CanUse)) Refresh();
    }
    private void ResetContext() {
        _generation++; _operation?.Cancel(); _source = null; _snapshotHash = null; _history.Clear(); Messages.Clear();
        Status = Text("ContextReset", "Conversation cleared. The next question uses the current document, scope and connection.");
        Refresh();
    }
    private string Text(string key, string fallback) => _localizer.GetOrDefault("Assistant." + key, fallback);
    public void Dispose() { if (_disposed) return; _disposed = true; Connections.Changed -= OnConnectionChanged; Connections.PropertyChanged -= OnConnectionPropertyChanged; ResetContext(); }
}

internal sealed record AssistantMessage(string Text, IReadOnlyList<AssistantCitation> Citations, bool IsQuestion);

internal sealed class AssistantCitation {
    internal AssistantCitation(OfficeAiCitation source, Action navigate, IStudioLocalizer localizer) {
        Label = source.Page is int page ? localizer.FormatOrDefault("Assistant.SourcePage", "Page {0} · {1}", page, source.EvidenceId) : source.EvidenceId;
        Quote = source.Quote ?? string.Empty;
        NavigateCommand = new RelayCommand(navigate, () => source.Page.HasValue);
    }
    public string Label { get; }
    public string Quote { get; }
    public IRelayCommand NavigateCommand { get; }
}
