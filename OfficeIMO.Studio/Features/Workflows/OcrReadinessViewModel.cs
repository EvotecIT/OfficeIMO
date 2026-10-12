using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>
/// Whether text recognition can run here, checked when an OCR screen opens: a build without external tools explains
/// that, and a missing Tesseract engine offers the way to set it up instead of failing only when recognition starts.
/// </summary>
public sealed partial class OcrReadinessViewModel : ObservableObject, IDisposable {
    private readonly StudioPreferencesService? _preferences;
    private readonly IStudioLocalizer _localizer;
    private readonly Func<string?, bool> _engineFound;

    internal OcrReadinessViewModel(StudioPreferencesService? preferences, IStudioLocalizer localizer, Func<string?, bool>? engineFound = null) {
        _preferences = preferences;
        _localizer = localizer;
        _engineFound = engineFound ?? (path => TesseractRuntime.TryDiscover(path, out _));
        if (_preferences is not null) _preferences.Changed += OnPreferencesChanged;
    }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasNotice))]
    private string? _notice;

    /// <summary>True when the engine is missing and Settings can fix it; false for a build that cannot run it at all.</summary>
    [ObservableProperty]
    private bool _canSetUp;

    public bool HasNotice => Notice is not null;

    internal void Refresh() {
        if (StudioOcrProvider.UnavailableReason is { } reason) {
            (Notice, CanSetUp) = (reason, false);
            return;
        }
        bool found = _engineFound(_preferences?.Current.OcrExecutablePath);
        (Notice, CanSetUp) = found ? (null, false) : (_localizer.Get("Ocr.Notice.EngineMissing"), true);
    }

    /// <summary>Raised when the user asks to set up recognition; the window shows Settings at text recognition.</summary>
    internal event EventHandler? SetUpRequested;

    [RelayCommand]
    private void SetUp() => SetUpRequested?.Invoke(this, EventArgs.Empty);

    private void OnPreferencesChanged(object? sender, EventArgs e) {
        if (HasNotice) Refresh();
    }

    public void Dispose() {
        if (_preferences is not null) _preferences.Changed -= OnPreferencesChanged;
    }
}
