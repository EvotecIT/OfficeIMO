using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Security;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal bool IsEncrypted => _documentInfo.Security.HasEncryption;

    internal bool CanChangeEncryption(string? ownerPassword) {
        try {
            PdfLoadOptions readOptions = IsEncrypted && !_documentInfo.Security.HasOwnerAuthorization
                ? new PdfLoadOptions { Password = ownerPassword }
                : _readOptions;
            return PdfDocument.Load(_bytes, readOptions).PlanMutation(PdfMutationOperation.ChangeEncryption).CanExecute;
        } catch {
            return false;
        }
    }

    internal bool CanSign => CanPlan(PdfMutationOperation.PrepareExternalSignature);

    internal Task SignAsync(
        X509Certificate2 certificate,
        PdfExternalSignatureOptions options,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        ArgumentNullException.ThrowIfNull(certificate);
        ArgumentNullException.ThrowIfNull(options);
        if (!CanSign) throw new InvalidOperationException("This document cannot accept another signature under its current security policy.");
        options.CancellationToken = cancellationToken;
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.Signature,
            "Applied certificate signature " + options.FieldName,
            options.VisibleAppearance is null ? Array.Empty<int>() : new[] { options.VisibleAppearance.PageNumber },
            bytes => {
                using var signer = new PdfCmsExternalSigner(
                    OfficeSecurityProvider.Default,
                    certificate,
                    string.IsNullOrWhiteSpace(options.Name) ? null : options.Name);
                return LoadDocument(bytes).Security.SignExternal(signer, options).Pdf;
            },
            cancellationToken,
            progress,
            detachCpuWorkOnCancellation: false);
    }

    internal Task<PdfSignatureValidationReport> ValidateSignaturesAsync(CancellationToken cancellationToken) {
        ThrowIfDisposed();
        byte[] snapshot = CopyBytes();
        return RunCancellableCpuWorkAsync(() => {
            var provider = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default);
            return LoadDocument(snapshot).Security.ValidateSignatures(provider);
        }, cancellationToken);
    }

    internal Task ApplyBatesNumberingAsync(
        PdfBatesNumberingOptions options,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        ArgumentNullException.ThrowIfNull(options);
        if (!CanEditPageContent) throw new InvalidOperationException("This document cannot safely add Bates numbers.");
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.BatesNumbering,
            "Applied Bates numbering",
            Enumerable.Range(1, Pages.Count).ToArray(),
            bytes => {
                var source = new PdfBatesDocument(bytes, FileName) { ReadOptions = _readOptions };
                return PdfBatesNumberer.Apply(new[] { source }, options).Documents[0].ToBytes();
            },
            cancellationToken,
            progress);
    }

    private string ValidateExportDestination(string destinationPath) {
        if (string.IsNullOrWhiteSpace(destinationPath)) throw new ArgumentException("Choose an output PDF.", nameof(destinationPath));
        string destination = OfficeIMO.Internal.OfficeStorageIdentity.Normalize(destinationPath);
        if (OfficeIMO.Internal.OfficeStorageIdentity.AreEquivalent(destination, Path)) {
            throw new InvalidOperationException("Choose a different output path so the open document remains unchanged.");
        }
        return destination;
    }

}
