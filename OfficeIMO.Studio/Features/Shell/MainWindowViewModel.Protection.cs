using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private readonly Func<PdfProtectionPreviewViewModel, Task<bool>> _reviewProtection;
    private readonly Func<PdfProtectionPreviewViewModel, Task> _showProtectionResult;
    private bool _reviewingProtection;

    [RelayCommand]
    private Task SaveProtectedCopyAsync(CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(ProtectUserPassword)) {
            ErrorMessage = UiText("Protection.PasswordRequired"); return Task.CompletedTask;
        }
        if (!string.Equals(ProtectUserPassword, ProtectConfirmPassword, StringComparison.Ordinal)) {
            ErrorMessage = UiText("Protection.PasswordMismatch"); return Task.CompletedTask;
        }
        var encryption = new PdfStandardEncryptionOptions(ProtectUserPassword) {
            OwnerPassword = string.IsNullOrWhiteSpace(ProtectOwnerPassword) ? null : ProtectOwnerPassword,
            EncryptMetadata = ProtectEncryptMetadata, AllowedPermissions = BuildProtectionPermissions()
        };
        return SaveProtectionCopyAsync(encryption, CurrentOwnerPassword, cancellationToken);
    }

    [RelayCommand]
    private Task SaveDecryptedCopyAsync(CancellationToken cancellationToken) =>
        SaveProtectionCopyAsync(null, CurrentOwnerPassword, cancellationToken);

    private async Task SaveProtectionCopyAsync(PdfStandardEncryptionOptions? encryption, string? currentOwnerPassword,
        CancellationToken cancellationToken) {
        if (_workspace is null || IsWorkspaceBusy || _reviewingProtection) return;
        var workspace = _workspace;
        if (encryption is null && !workspace.IsEncrypted) return;
        if (!workspace.CanChangeEncryption(currentOwnerPassword)) {
            ErrorMessage = UiText("Capability.ProtectionUnavailable"); return;
        }
        long revision = workspace.Revision;
        if (!IsReviewedCopyCurrent(workspace, revision)) return;
        _reviewingProtection = true;
        try {
            string? destination = await _pickSavePdf(cancellationToken).ConfigureAwait(true);
            if (string.IsNullOrWhiteSpace(destination) || !IsReviewedCopyCurrent(workspace, revision)) return;
            bool provider = _services.Storage.UsesProviderPublication(destination);
            var preview = new PdfProtectionPreviewViewModel(workspace.Pages.Count, destination, provider, encryption, _localizer,
                path => _openDocumentInTab is null ? _openUri(new Uri(path)) : _openDocumentInTab(path, CancellationToken.None),
                path => _openUri(new Uri(Path.GetDirectoryName(path)!)));
            if (!await _reviewProtection(preview).ConfigureAwait(true) || !IsReviewedCopyCurrent(workspace, revision)) return;
            if (provider && !await _confirmProviderWrite(destination).ConfigureAwait(true)) return;
            if (!IsReviewedCopyCurrent(workspace, revision)) return;
            OfficeWorkflowResult? result = null;
            await RunStandaloneAsync(async token => {
                StudioJobRecord? job = null;
                try {
                    var output = _services.Storage.CreateWorkflowOutput(destination, _services.WorkflowRecovery);
                    job = _services.Jobs.Start(preview.Title, workspace.Path, destination, CancelCurrentOperation);
                    using IDisposable execution = await _services.Jobs.EnterAsync(token).ConfigureAwait(true);
                    var progress = new Progress<Features.Workspace.PdfWorkspaceProgress>(update => {
                        if (!job.IsActive) return;
                        OperationStatus = update.Stage; OperationProgressFraction = update.Fraction;
                        job.Report(new OfficeWorkflowProgress("protection", "protection", update.Stage, update.Fraction));
                    });
                    if (!IsReviewedCopyCurrent(workspace, revision))
                        throw new InvalidOperationException(ErrorMessage ?? UiText("Organizer.StalePreview"));
                    result = encryption is null
                        ? await workspace.SaveDecryptedCopyAsync(destination, currentOwnerPassword, token, progress, output, _publicationGuard).ConfigureAwait(true)
                        : await workspace.SaveProtectedCopyAsync(destination, encryption, currentOwnerPassword, token, progress, output, _publicationGuard).ConfigureAwait(true);
                    job.Complete(result.Status, result.OutputPath, result.Summary, result.Recovery);
                } catch (OperationCanceledException) when (token.IsCancellationRequested) {
                    job?.Complete(OfficeWorkflowStatus.Cancelled, null, UiText("Workspace.OperationCancelled")); throw;
                } catch (Exception error) {
                    job?.Complete(OfficeWorkflowStatus.Failed, null, error.Message); throw;
                }
            }, cancellationToken).ConfigureAwait(true);
            if (result is not null && !_disposed && ReferenceEquals(workspace, _workspace)) {
                OperationStatus = result.Summary;
                if (result.Status is OfficeWorkflowStatus.Failed or OfficeWorkflowStatus.Unconfirmed) ErrorMessage = result.Summary;
                if (result.Succeeded) {
                    // Do not erase newly entered settings while a captured request was running.
                    if (encryption is not null) {
                        if (ProtectUserPassword == encryption.UserPassword) ProtectUserPassword = string.Empty;
                        if (ProtectConfirmPassword == encryption.UserPassword) ProtectConfirmPassword = string.Empty;
                        if (ProtectOwnerPassword == encryption.OwnerPassword) ProtectOwnerPassword = string.Empty;
                    }
                    if (CurrentOwnerPassword == currentOwnerPassword) CurrentOwnerPassword = string.Empty;
                }
                preview.Complete(result);
                await _showProtectionResult(preview).ConfigureAwait(true);
            }
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            OperationStatus = UiText("Workspace.OperationCancelled");
        } catch (Exception error) { ErrorMessage = error.Message; }
        finally { _reviewingProtection = false; }
    }
}
