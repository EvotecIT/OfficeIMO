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
