using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Settings;

/// <summary>Checks OCR prerequisites and persists only a verified executable choice.</summary>
internal sealed partial class StudioOcrSetupViewModel : ObservableObject, IDisposable {
    private readonly StudioPreferencesService _preferences;
    private readonly IStudioLocalizer _localizer;
    private readonly IStudioDiagnostics _diagnostics;
    private readonly Func<CancellationToken, Task<string?>> _pickExecutable;
    private readonly Func<string?, CancellationToken, Task<StudioOcrRuntimeInfo>> _inspect;
    private readonly CancellationTokenSource _lifetime = new();
    private string _savedPath;
    private bool _disposed;

    internal StudioOcrSetupViewModel(StudioPreferencesService preferences, IStudioLocalizer localizer,
        IStudioDiagnostics diagnostics, Func<CancellationToken, Task<string?>>? pickExecutable = null,
        StudioOcrRuntime? runtime = null,
        Func<string?, CancellationToken, Task<StudioOcrRuntimeInfo>>? inspect = null) {
        _preferences = preferences; _localizer = localizer; _diagnostics = diagnostics;
        _pickExecutable = pickExecutable ?? (_ => Task.FromResult<string?>(null));
        _inspect = inspect ?? (runtime ?? new StudioOcrRuntime(preferences)).InspectAsync;
        _executablePath = preferences.Current.OcrExecutablePath ?? string.Empty;
        _savedPath = _executablePath;
        _status = StudioOcrProvider.UnavailableReason ?? T("CheckToStart");
        _preferences.Changed += PreferencesChanged;
    }

    public bool IsAvailable => StudioOcrProvider.UnavailableReason is null;
    public string InstallationCommand => TesseractRuntime.GetInstallationHint();
    public string LanguageHelp => T("LanguageHelp");
    public bool CanCheck => IsAvailable && !_disposed && !IsBusy;

    [ObservableProperty] private string _executablePath;
    [ObservableProperty] private string _status;
    [ObservableProperty] private string? _detectedExecutable;
    [ObservableProperty] private string? _version;
    [ObservableProperty] private string? _installedLanguages;
    [ObservableProperty] private bool _isReady;
    [ObservableProperty] private bool _isBusy;
    partial void OnIsBusyChanged(bool value) => NotifyCommands();
    partial void OnExecutablePathChanged(string value) {
        IsReady = false; DetectedExecutable = null; Version = null; InstalledLanguages = null;
        if (!IsBusy) Status = T("CheckToStart");
    }

    [RelayCommand(CanExecute = nameof(CanCheck))]
    private async Task CheckAsync() => await CheckAndSaveAsync(ExecutablePath);

    [RelayCommand(CanExecute = nameof(CanCheck))]
    private async Task ChooseExecutableAsync() {
        try {
            string? path = await _pickExecutable(_lifetime.Token);
            if (_disposed || string.IsNullOrWhiteSpace(path)) return;
            ExecutablePath = path;
            await CheckAndSaveAsync(path);
        } catch (OperationCanceledException) when (_lifetime.IsCancellationRequested) { }
    }

    [RelayCommand(CanExecute = nameof(CanCheck))]
    private async Task UseAutomaticDiscoveryAsync() {
        ExecutablePath = string.Empty;
        await CheckAndSaveAsync(null);
    }

    private async Task CheckAndSaveAsync(string? path) {
        if (!CanCheck) return;
        IsBusy = true; IsReady = false; Status = T("Checking");
        DetectedExecutable = null; Version = null; InstalledLanguages = null;
        try {
            string? selected = string.IsNullOrWhiteSpace(path) ? null : path.Trim();
            if (selected is not null && !Path.IsPathFullyQualified(selected))
                throw new ArgumentException(T("AbsolutePathRequired"));
            StudioOcrRuntimeInfo result = await _inspect(selected, _lifetime.Token);
            if (_disposed) return;
            _preferences.Update(current => current with { OcrExecutablePath = selected is null ? null : result.ExecutablePath });
            DetectedExecutable = result.ExecutablePath;
            Version = result.Version;
            InstalledLanguages = result.Languages.Count == 0 ? T("NoLanguages") : string.Join(", ", result.Languages);
            IsReady = true;
            Status = T("Ready");
        } catch (OperationCanceledException) when (_lifetime.IsCancellationRequested) {
            // Closing Settings cancels the external prerequisite probe.
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException or ArgumentException
            or InvalidOperationException or TimeoutException or NotSupportedException or System.ComponentModel.Win32Exception) {
            if (_disposed) return;
            _diagnostics.Write(StudioDiagnosticLevel.Warning, "Ocr", "RuntimeSetupFailed", error);
            Status = error is FileNotFoundException ? T("Missing") : _localizer.Format("OcrSetup.CheckFailed", error.Message);
        } finally {
            if (!_disposed) IsBusy = false;
        }
    }

    private void NotifyCommands() {
        OnPropertyChanged(nameof(CanCheck));
        CheckCommand.NotifyCanExecuteChanged(); ChooseExecutableCommand.NotifyCanExecuteChanged();
        UseAutomaticDiscoveryCommand.NotifyCanExecuteChanged();
    }
    private string T(string name) => _localizer.Get("OcrSetup." + name);
    private void PreferencesChanged(object? sender, EventArgs args) {
        string next = _preferences.Current.OcrExecutablePath ?? string.Empty;
        if (next == _savedPath) return;
        bool hasDraft = ExecutablePath != _savedPath;
        _savedPath = next;
        if (!IsBusy && !hasDraft) ExecutablePath = next;
    }
    public void Dispose() {
        if (_disposed) return;
        _disposed = true; _preferences.Changed -= PreferencesChanged;
        _lifetime.Cancel(); _lifetime.Dispose(); NotifyCommands();
    }
}
