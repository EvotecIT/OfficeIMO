using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Provenance.C2pa;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ProvenanceWorkbenchViewModel {
    /// <summary>External executables are configured only by desktop hosts with local paths.</summary>
    public bool CanConfigureProvider => !UsesWorkingCopies && !OperatingSystem.IsAndroid() && !OperatingSystem.IsIOS() && !OperatingSystem.IsBrowser();
    [ObservableProperty] private string _c2paToolPath = "";
    [ObservableProperty] private string _trustAnchorsPath = "";
    [ObservableProperty] private string _allowedListPath = "";
    [ObservableProperty] private string _providerStatus = "No verifier configured. Assessment remains available without it.";
    partial void OnC2paToolPathChanged(string value) => InvalidateProvider();
    partial void OnTrustAnchorsPathChanged(string value) => InvalidateProvider();
    partial void OnAllowedListPathChanged(string value) => InvalidateProvider();
    private void InvalidateProvider() {
        _cancellation?.Cancel();
        OnInputPathChanged(InputPath);
        ProviderStatus = "Configuration changed. Check the executable, then assess the file again.";
        OnPropertyChanged(nameof(CanCreateCopy));
    }
    [RelayCommand] private async Task CheckProviderAsync() {
        if (_disposed || IsBusy || !CanConfigureProvider) return;
        if (string.IsNullOrWhiteSpace(C2paToolPath)) { ProviderStatus = "Enter the path to a trusted c2patool executable."; return; }
        string executable = C2paToolPath.Trim();
        int revision = _revision;
        using var cancellation = new CancellationTokenSource();
        _cancellation = cancellation; IsBusy = true;
        ProviderStatus = "Checking executable and process containment…";
        try {
            C2paToolAvailability result = await Task.Run(() => new C2paToolProvenanceVerifier(executable)
                .CheckAvailability(cancellationToken: cancellation.Token), cancellation.Token);
            if (!_disposed && revision == _revision)
                ProviderStatus = $"{(result.Available ? "Available" : "Unavailable")} · {result.Version ?? "unknown version"} · {result.Diagnostic}";
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            if (!_disposed && revision == _revision) ProviderStatus = "Provider check cancelled.";
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            if (!_disposed && revision == _revision) ProviderStatus = "Provider check failed: " + error.Message;
        } finally { _cancellation = null; IsBusy = false; }
    }
}
