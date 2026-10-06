using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed class StudioSessionEnvironment : IStudioSessionEnvironment {
    private readonly StudioApplicationServices _services;
    internal StudioSessionEnvironment(StudioApplicationServices services) {
        _services = services;
        _services.DocumentHistory.Cleared += OnHistoryCleared;
    }
    public StudioSessionStore SessionStore => _services.DocumentHistory.RestartSession;
    public PdfWorkspaceRecoveryStore Recovery => _services.Recovery;
    public StudioDocumentStorage Storage => _services.Storage;
    public bool RememberSession => _services.Preferences.Current.RememberSession;
    public event EventHandler? PreferencesChanged {
        add => _services.Preferences.Changed += value;
        remove => _services.Preferences.Changed -= value;
    }
    public event EventHandler? SessionCleared;
    private void OnHistoryCleared(object? sender, StudioHistoryCleanupResult result) {
        if (result.RestartSession) SessionCleared?.Invoke(this, EventArgs.Empty);
    }
    public string Text(string key) => key == "SummaryOne"
        ? _services.Localizer.GetOrDefault("Session.SummaryOne", "1 document from your last session")
        : _services.Localizer.Get("Session." + key);
    public string Format(string key, int count) =>
        _services.Localizer.FormatOrDefault("Session." + key, "{0:N0} documents from your last session", count);
    public void ReportFailure(string code, Exception error) => _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "Session", code, error);
    public void Dispose() => _services.DocumentHistory.Cleared -= OnHistoryCleared;
}
