using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure.Diagnostics;

namespace OfficeIMO.Studio.Features.Settings;

internal sealed partial class StudioSettingsViewModel {
    [ObservableProperty] private bool _confirmHistoryClear;
    [ObservableProperty] private string? _historyStatus;
    public bool HasHistoryStatus => !string.IsNullOrWhiteSpace(HistoryStatus);
    partial void OnHistoryStatusChanged(string? value) => OnPropertyChanged(nameof(HasHistoryStatus));
    public bool RememberDocumentHistory => _history.RememberHistory;
    public string HistoryPersistenceAction => _localizer.Get(RememberDocumentHistory
        ? "Settings.DisableHistory" : "Settings.EnableHistory");

    [RelayCommand]
    private void ToggleDocumentHistory() {
        try {
            _history.SetRememberHistory(!RememberDocumentHistory);
            HistoryStatus = null;
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            _diagnostics.Write(StudioDiagnosticLevel.Warning, "Preferences", "HistoryPrivacyChangeFailed", error);
            HistoryStatus = _localizer.Get("Settings.HistoryPreferenceFailed");
        }
    }

    /// <summary>Previews render the first page of recent local documents in memory; private work can turn them off.</summary>
    public bool ShowRecentPreviews {
        get => _preferences.Current.ShowRecentPreviews;
        set {
            if (value == ShowRecentPreviews) return;
            try {
                _preferences.Update(current => current with { ShowRecentPreviews = value });
                HistoryStatus = null;
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                _diagnostics.Write(StudioDiagnosticLevel.Warning, "Preferences", "PreviewPrivacyChangeFailed", error);
                HistoryStatus = _localizer.Get("Settings.HistoryPreferenceFailed");
            }
            OnPropertyChanged();
        }
    }
    [RelayCommand]
    private void RequestHistoryClear() { HistoryStatus = null; ConfirmHistoryClear = true; }

    [RelayCommand]
    private void CancelHistoryClear() => ConfirmHistoryClear = false;

    [RelayCommand]
    private void ClearHistory() {
        if (!ConfirmHistoryClear) return;
        var result = _history.Clear();
        HistoryStatus = _localizer.Get(result.Succeeded ? "Settings.HistoryCleared" : "Settings.HistoryClearFailed");
        ConfirmHistoryClear = false;
    }
}
