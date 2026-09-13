using System.Security.Cryptography.X509Certificates;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private readonly Func<PdfSigningPreviewViewModel, Task<bool>> _reviewSigning;
    private readonly Func<PdfSigningPreviewViewModel, Task> _showSigningResult;
    private readonly Func<string, X509Certificate2> _loadSigningCertificate;
    private bool _reviewingSigning;

    [RelayCommand]
    private async Task ApplyCertificateSignatureAsync(CancellationToken cancellationToken) {
        using var notifications = BeginNotificationScope();
        if (_workspace is null || SelectedSigningCertificate is null || IsWorkspaceBusy || _reviewingSigning) return;
        var workspace = _workspace; long revision = workspace.Revision;
        if (!IsReviewedCopyCurrent(workspace, revision)) return;
        if (!workspace.CanSign) { ErrorMessage = UiText("Signing.Unavailable"); return; }
        var settings = new PdfSigningSettings(SelectedSigningCertificate, SignatureFieldName, SignatureReason, SignatureLocation,
            SignatureIsVisible, SignaturePageNumber, SignatureX, SignatureY, SignatureWidth, SignatureHeight);
        _reviewingSigning = true;
        try {
            string? destination = await _pickSavePdf(cancellationToken).ConfigureAwait(true);
            if (string.IsNullOrWhiteSpace(destination) || !IsReviewedCopyCurrent(workspace, revision)) return;
            byte[]? image = null;
            if (!await RunStandaloneAsync(async token => image = await workspace.PreviewSignatureAsync(settings.CreateOptions(), token).ConfigureAwait(true),
                    cancellationToken).ConfigureAwait(true) || !IsReviewedCopyCurrent(workspace, revision)) return;
            bool provider = _services.Storage.UsesProviderPublication(destination);
            using var preview = new PdfSigningPreviewViewModel(settings, workspace.Pages.Count, destination, provider, image, _localizer,
                path => _openDocumentInTab is null ? _openUri(new Uri(path)) : _openDocumentInTab(path, CancellationToken.None),
                path => _openUri(new Uri(Path.GetDirectoryName(path)!)));
            if (!await _reviewSigning(preview).ConfigureAwait(true) || !IsReviewedCopyCurrent(workspace, revision)) return;
            if (provider && !await _confirmProviderWrite(destination).ConfigureAwait(true)) return;
            if (!IsReviewedCopyCurrent(workspace, revision)) return;
            OfficeWorkflowResult? result = null;
            await RunStandaloneAsync(async token => {
                StudioJobRecord? job = null;
                try {
                    job = _services.Jobs.Start(preview.Title, workspace.Path, destination, CancelCurrentOperation);
                    using IDisposable execution = await _services.Jobs.EnterAsync(token).ConfigureAwait(true);
                    if (!IsReviewedCopyCurrent(workspace, revision)) throw new InvalidOperationException(ErrorMessage ?? UiText("Organizer.StalePreview"));
                    using X509Certificate2 certificate = _loadSigningCertificate(settings.Certificate.Thumbprint);
                    if (!string.Equals(certificate.Thumbprint, settings.Certificate.Thumbprint, StringComparison.OrdinalIgnoreCase) ||
                        !certificate.HasPrivateKey || certificate.NotBefore > DateTime.Now || certificate.NotAfter < DateTime.Now)
                        throw new InvalidOperationException(UiText("Signing.CertificateChanged"));
                    using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
                    var verifier = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default);
                    var progress = new Progress<Features.Workspace.PdfWorkspaceProgress>(update => {
                        if (!job.IsActive) return;
                        OperationStatus = update.Stage; OperationProgressFraction = update.Fraction;
                        job.Report(new OfficeWorkflowProgress("signing", "signing", update.Stage, update.Fraction));
                    });
                    var output = _services.Storage.CreateWorkflowOutput(destination, _services.WorkflowRecovery);
                    result = await workspace.SaveSignedCopyAsync(destination, signer, settings.CreateOptions(), verifier, token, progress,
                        output, _publicationGuard).ConfigureAwait(true);
                    job.Complete(result.Status, result.OutputPath, result.Summary, result.Recovery);
                } catch (OperationCanceledException) when (token.IsCancellationRequested) {
                    job?.Complete(OfficeWorkflowStatus.Cancelled, null, UiText("Workspace.OperationCancelled")); throw;
                } catch (Exception error) { job?.Complete(OfficeWorkflowStatus.Failed, null, error.Message); throw; }
            }, cancellationToken).ConfigureAwait(true);
            if (result is not null && !_disposed && ReferenceEquals(workspace, _workspace)) {
                OperationStatus = result.Summary;
                if (result.Status is OfficeWorkflowStatus.Failed or OfficeWorkflowStatus.Unconfirmed) ErrorMessage = result.Summary;
                preview.Complete(result); await _showSigningResult(preview).ConfigureAwait(true);
            }
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { OperationStatus = UiText("Workspace.OperationCancelled"); }
        catch (Exception error) { ErrorMessage = error.Message; }
        finally { _reviewingSigning = false; }
    }
}
