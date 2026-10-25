using OfficeIMO.Studio.Features.Home;

namespace OfficeIMO.Studio.Infrastructure.Preferences;

internal sealed record StudioHistoryCleanupResult(bool RecentDocuments, bool ReadingPositions, bool RestartSession) {
    internal bool ConversionRuns { get; init; } = true;
    internal bool Succeeded => RecentDocuments && ReadingPositions && RestartSession && ConversionRuns;
}

/// <summary>Owns UI-thread document-history policy and explicit cleanup across its stores.</summary>
internal sealed class StudioDocumentHistory {
    private readonly StudioPreferencesService _preferences;

    internal StudioDocumentHistory(StudioDataPaths paths, StudioPreferencesService preferences, StudioLocalDocumentRoot? documents = null) {
        _preferences = preferences;
        RecentDocuments = new(paths.RecentDocumentsPath, () => RememberHistory, documents);
        ReadingPositions = new(paths.DocumentViewsPath, () => RememberHistory, documents);
        RestartSession = new(paths.SessionPath);
        ConversionRuns = new(paths.ConversionRunsRoot, () => RememberHistory);
    }

    internal bool RememberHistory => _preferences.Current.RememberDocumentHistory;
    internal JsonRecentDocumentStore RecentDocuments { get; }
    internal StudioDocumentViewStore ReadingPositions { get; }
    internal StudioSessionStore RestartSession { get; }
    internal StudioConversionRunStore ConversionRuns { get; }
    internal event EventHandler<StudioHistoryCleanupResult>? Cleared;

    /// <summary>Raised after any window or tab rewrites the recent-document list; the sender is the writer.</summary>
    internal event EventHandler? RecentDocumentsChanged;

    internal void NotifyRecentDocumentsChanged(object writer) => RecentDocumentsChanged?.Invoke(writer, EventArgs.Empty);

    /// <summary>Clears only the recent-document list and lets every open tab drop its copy.</summary>
    internal void ClearRecentDocuments(object writer) {
        RecentDocuments.Clear();
        NotifyRecentDocumentsChanged(writer);
    }

    internal void SetRememberHistory(bool enabled) =>
        _preferences.Update(current => current with { RememberDocumentHistory = enabled });

    internal void SetRememberSession(bool enabled) {
        _preferences.Update(current => current with { RememberSession = enabled });
        if (!enabled) RestartSession.Clear();
    }

    internal StudioHistoryCleanupResult Clear() {
        // Attempt every store even if one file is locked. Report partial removal truthfully.
        var result = new StudioHistoryCleanupResult(
            TryClear(RecentDocuments.Clear), TryClear(ReadingPositions.Clear), TryClear(RestartSession.Clear)) {
            ConversionRuns = ConversionRuns.ClearFinished()
        };
        Cleared?.Invoke(this, result);
        return result;
    }

    private static bool TryClear(Action clear) {
        try { clear(); return true; }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { return false; }
    }
}
