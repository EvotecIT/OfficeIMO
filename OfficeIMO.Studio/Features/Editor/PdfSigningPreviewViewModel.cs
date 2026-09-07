using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class PdfSigningPreviewViewModel : ObservableObject, IDisposable {
    private readonly IStudioLocalizer _localizer;
    private readonly Func<string, Task> _open;
    private readonly Func<string, Task> _reveal;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Title))]
    [NotifyPropertyChangedFor(nameof(IsPreview))]
    private bool _hasResult;
    [ObservableProperty] private string? _summary;
    [ObservableProperty] private string? _errorMessage;
    [ObservableProperty] private string? _outputPath;
    [ObservableProperty] private bool _hasRecovery;
    [ObservableProperty] private string? _verification;
    [ObservableProperty] private IReadOnlyList<string> _findings = [];
    internal PdfSigningPreviewViewModel(PdfSigningSettings settings, int pages, string destination, bool provider,
        byte[]? image, IStudioLocalizer localizer, Func<string, Task> open, Func<string, Task> reveal) {
        _localizer = localizer; _open = open; _reveal = reveal; Destination = destination;
        Details = [localizer.Format("Protection.Pages", pages),
            localizer.Format("Signing.Certificate", settings.Certificate.DisplayName),
            localizer.Format("Signing.Thumbprint", settings.Certificate.Thumbprint),
            localizer.Format("Signing.Issuer", settings.Certificate.Issuer),
            localizer.Format("Signing.Expires", settings.Certificate.NotAfter.ToString("yyyy-MM-dd")),
            localizer.Format("Signing.Field", settings.FieldName),
            localizer.Format("Signing.Reason", settings.Reason), localizer.Format("Signing.Location", settings.Location),
            localizer.Get(settings.Visible ? "Signing.VisibleHint" : "Signing.InvisibleHint"),
            localizer.Get(provider ? "Protection.ProviderHint" : "Protection.LocalHint"), localizer.Get("Signing.TrustPolicy")];
        if (image is not null) { using var stream = new MemoryStream(image, false); PreviewImage = new Bitmap(stream); }
    }
    public string Title => _localizer.Get(HasResult ? "Signing.ResultTitle" : "Signing.Title");
    public string Destination { get; }
    public IReadOnlyList<string> Details { get; }
    public Bitmap? PreviewImage { get; }
    public bool HasPreviewImage => PreviewImage is not null;
    public bool IsPreview => !HasResult;
    public bool CanOpenOutput => HasResult && OutputPath is not null;
    public bool CanRevealOutput => CanOpenOutput && OfficeStorageIdentity.GetLocalPath(OutputPath!) is not null;
    internal void Complete(OfficeWorkflowResult result) {
        Summary = result.Summary; OutputPath = result.Succeeded ? result.OutputPath : null;
        HasRecovery = result.Recovery is not null; HasResult = true;
        if (result.SignatureReport is { } report) {
            Findings = report.Findings.Where(finding => finding.Severity != OfficeIMO.Pdf.PdfDiagnosticSeverity.Info)
                .Select(finding => finding.Severity + " · " + finding.Message).ToArray();
            Verification = string.Join(Environment.NewLine, new[] {
                _localizer.Get(result.Succeeded ? "Signing.PublishedEvidence" : "Signing.PreparedEvidence"),
                _localizer.Format("Signing.Math", _localizer.Get(report.MathematicalSignaturesVerified && report.DigestVerified ? "Signing.Verified" : "Signing.Unverified")),
                _localizer.Format("Signing.Chain", _localizer.Get(report.CertificateChainVerified ? "Signing.Verified" : "Signing.Unverified")),
                _localizer.Format("Signing.Revocation", _localizer.Get(report.RevocationChecked ? "Signing.Checked" : "Signing.NotChecked")),
                _localizer.Format("Signing.Timestamp", _localizer.Get(report.TimestampValidationPerformed ? "Signing.Checked" : "Signing.NotChecked"))
            });
        }
        OnPropertyChanged(nameof(CanOpenOutput)); OnPropertyChanged(nameof(CanRevealOutput));
    }
    [RelayCommand] private async Task OpenOutputAsync() {
        if (!CanOpenOutput) return;
        try { await _open(OutputPath!); } catch (Exception error) { ErrorMessage = error.Message; }
    }
    [RelayCommand] private async Task RevealOutputAsync() {
        if (!CanRevealOutput) return;
        try { await _reveal(OutputPath!); } catch (Exception error) { ErrorMessage = error.Message; }
    }
    public void Dispose() => PreviewImage?.Dispose();
}
