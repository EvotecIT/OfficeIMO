using OfficeIMO.Studio.Features.Home;

namespace OfficeIMO.Studio.Infrastructure.Preferences;

internal sealed record StudioHistoryCleanupResult(bool RecentDocuments, bool ReadingPositions, bool RestartSession) {
    internal bool Succeeded => RecentDocuments && ReadingPositions && RestartSession;
}

/// <summary>Owns UI-thread document-history policy and explicit cleanup across its three stores.</summary>
internal sealed class StudioDocumentHistory {
    private readonly StudioPreferencesService _preferences;

    internal StudioDocumentHistory(StudioDataPaths paths, StudioPreferencesService preferences) {
        _preferences = preferences;
        RecentDocuments = new(paths.RecentDocumentsPath, () => RememberHistory);
        ReadingPositions = new(paths.DocumentViewsPath, () => RememberHistory);
        RestartSession = new(paths.SessionPath);
    }

    internal bool RememberHistory => _preferences.Current.RememberDocumentHistory;
    internal JsonRecentDocumentStore RecentDocuments { get; }
    internal StudioDocumentViewStore ReadingPositions { get; }
    internal StudioSessionStore RestartSession { get; }
    internal event EventHandler<StudioHistoryCleanupResult>? Cleared;

    internal void SetRememberHistory(bool enabled) =>
        _preferences.Update(current => current with { RememberDocumentHistory = enabled });

    internal void SetRememberSession(bool enabled) {
        _preferences.Update(current => current with { RememberSession = enabled });
        if (!enabled) RestartSession.Clear();
    }

    internal StudioHistoryCleanupResult Clear() {
        // Attempt every store even if one file is locked. Report partial removal truthfully.
        var result = new StudioHistoryCleanupResult(
            TryClear(RecentDocuments.Clear), TryClear(ReadingPositions.Clear), TryClear(RestartSession.Clear));
        Cleared?.Invoke(this, result);
        return result;
    }

    private static bool TryClear(Action clear) {
        try { clear(); return true; }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { return false; }
    }
}
