using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class PdfProtectionPreviewViewModel : ObservableObject {
    private readonly IStudioLocalizer _localizer;
    private readonly Func<string, Task> _open;
    private readonly Func<string, Task> _reveal;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Title))]
    private bool _hasResult;
    [ObservableProperty] private string? _summary;
    [ObservableProperty] private string? _errorMessage;
    [ObservableProperty] private string? _outputPath;
    [ObservableProperty] private bool _hasRecovery;
    [ObservableProperty] private string? _verification;
    internal PdfProtectionPreviewViewModel(int pageCount, string destination, bool provider,
        PdfStandardEncryptionOptions? encryption, IStudioLocalizer localizer, Func<string, Task> open, Func<string, Task> reveal) {
        Destination = destination; _localizer = localizer; _open = open; _reveal = reveal;
        ReviewTitle = localizer.Get(encryption is null ? "Protection.RemoveTitle" : "Protection.Title");
        DestinationHint = localizer.Get(provider ? "Protection.ProviderHint" : "Protection.LocalHint");
        var details = new List<string> { localizer.Format("Protection.Pages", pageCount),
            localizer.Get(encryption is null ? "Protection.RemoveHint" : "Protection.EncryptionHint") };
        if (encryption is not null) {
            details.Add(localizer.Get(encryption.EncryptMetadata ? "Protection.MetadataEncrypted" : "Protection.MetadataReadable"));
            foreach (var (permission, key) in new[] {
                (PdfStandardPermissions.Print, "DocumentWorkspace.Print"),
                (PdfStandardPermissions.HighQualityPrint, "DocumentWorkspace.HighQualityPrint"),
                (PdfStandardPermissions.CopyContents, "DocumentWorkspace.CopyContent"),
                (PdfStandardPermissions.Accessibility, "DocumentWorkspace.Accessibility"),
                (PdfStandardPermissions.ModifyAnnotations, "DocumentWorkspace.Annotations"),
                (PdfStandardPermissions.FillForms, "DocumentWorkspace.FillForms"),
                (PdfStandardPermissions.AssembleDocument, "DocumentWorkspace.AssemblePages"),
                (PdfStandardPermissions.ModifyContents, "DocumentWorkspace.ChangeContent") })
                details.Add(localizer.Format(encryption.AllowedPermissions.HasFlag(permission) ? "Protection.Allowed" : "Protection.Restricted", localizer.Get(key)));
            if (encryption.OwnerPassword is null || encryption.OwnerPassword == encryption.UserPassword)
                details.Add(localizer.Get("Protection.SharedPassword"));
            if (!encryption.AllowedPermissions.HasFlag(PdfStandardPermissions.CopyContents))
                details.Add(localizer.Get("Protection.StudioOwnerRequired"));
            details.Add(localizer.Get("Protection.ReaderPermissions"));
        }
        Details = details;
    }
    public string Title => HasResult ? _localizer.Get("Protection.ResultTitle") : ReviewTitle;
    private string ReviewTitle { get; }
    public string Destination { get; }
    public string DestinationHint { get; }
    public IReadOnlyList<string> Details { get; }
    public bool IsPreview => !HasResult;
    public bool CanOpenOutput => HasResult && OutputPath is not null;
    public bool CanRevealOutput => CanOpenOutput && OfficeStorageIdentity.GetLocalPath(OutputPath!) is not null;
    internal void Complete(OfficeWorkflowResult result) {
        Summary = result.Summary; OutputPath = result.Succeeded ? result.OutputPath : null;
        HasRecovery = result.Recovery is not null; HasResult = true;
        Verification = result.HealthReport is { Verified: true, After: not null } report
            ? _localizer.Format(result.Succeeded ? "Protection.Verified" : "Protection.VerifiedPrepared", report.Before.PageCount, report.After.PageCount, report.After.SizeBytes) : null;
        ErrorMessage = result.Succeeded ? null : string.Join(Environment.NewLine, result.Diagnostics
            .Where(item => item.Severity != OfficeWorkflowDiagnosticSeverity.Information && item.Message != result.Summary)
            .Select(item => item.Message).Distinct());
        OnPropertyChanged(nameof(IsPreview)); OnPropertyChanged(nameof(CanOpenOutput)); OnPropertyChanged(nameof(CanRevealOutput));
        OpenOutputCommand.NotifyCanExecuteChanged(); RevealOutputCommand.NotifyCanExecuteChanged();
    }
    [RelayCommand(CanExecute = nameof(CanOpenOutput))]
    private async Task OpenOutputAsync() {
        if (!CanOpenOutput) return;
        try { await _open(OutputPath!).ConfigureAwait(true); } catch (Exception error) { ErrorMessage = error.Message; }
    }
    [RelayCommand(CanExecute = nameof(CanRevealOutput))]
    private async Task RevealOutputAsync() {
        if (!CanRevealOutput) return;
        try { await _reveal(OfficeStorageIdentity.GetLocalPath(OutputPath!)!).ConfigureAwait(true); }
        catch (Exception error) { ErrorMessage = error.Message; }
    }
}
