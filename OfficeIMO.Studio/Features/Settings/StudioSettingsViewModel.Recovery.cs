using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure.Diagnostics;

namespace OfficeIMO.Studio.Features.Settings;

internal sealed partial class StudioSettingsViewModel {
    [ObservableProperty] private bool _confirmRecoveryClear;
    [ObservableProperty] private bool _isRecoveryBusy;
    [ObservableProperty] private string? _recoveryStatus;
    public bool HasRecoveryStatus => !string.IsNullOrWhiteSpace(RecoveryStatus);
    partial void OnRecoveryStatusChanged(string? value) => OnPropertyChanged(nameof(HasRecoveryStatus));

    [RelayCommand]
    private void RequestRecoveryClear() {
        if (IsRecoveryBusy) return;
        RecoveryStatus = null;
        ConfirmRecoveryClear = true;
    }

    [RelayCommand]
    private void CancelRecoveryClear() {
        if (!IsRecoveryBusy) ConfirmRecoveryClear = false;
    }

    [RelayCommand]
    private async Task ClearRecoveryAsync(CancellationToken token) {
        if (!ConfirmRecoveryClear || IsRecoveryBusy) return;
        IsRecoveryBusy = true;
        try {
            var result = await _recovery.ClearAllAsync(token);
            RecoveryStatus = result.FailedFiles == 0
                ? _localizer.Format("Settings.RecoveryCleared", result.RemovedFiles)
                : _localizer.Format("Settings.RecoveryClearIncomplete", result.RemovedFiles, result.FailedFiles);
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            _diagnostics.Write(StudioDiagnosticLevel.Warning, "Recovery", "ExplicitCleanupFailed", error);
            RecoveryStatus = _localizer.Get("Settings.RecoveryClearFailed");
        } catch (OperationCanceledException) when (token.IsCancellationRequested) {
            RecoveryStatus = _localizer.Get("Settings.RecoveryClearCanceled");
        } finally {
            IsRecoveryBusy = false;
            ConfirmRecoveryClear = false;
        }
    }
}
