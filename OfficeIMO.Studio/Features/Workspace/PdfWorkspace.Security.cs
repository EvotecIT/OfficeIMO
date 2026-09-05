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

    internal async Task SaveProtectedCopyAsync(
        string destinationPath,
        PdfStandardEncryptionOptions encryption,
        string? currentOwnerPassword,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        ArgumentNullException.ThrowIfNull(encryption);
        if (!CanChangeEncryption(currentOwnerPassword)) {
            throw new InvalidOperationException("This document's signature, certification, usage-rights, or authorization policy prevents changing password protection.");
        }
        string destination = ValidateExportDestination(destinationPath);
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            progress?.Report(new PdfWorkspaceProgress("Applying password protection", 0.15D));
            byte[] output = await RunCancellableCpuWorkAsync(() => {
                PdfDocument document = LoadDocument(_bytes);
                if (!IsEncrypted) return document.Security.Encrypt(encryption).Pdf;
                string ownerPassword = _documentInfo.Security.HasOwnerAuthorization
                    ? _readOptions.Password ?? string.Empty
                    : currentOwnerPassword ?? string.Empty;
                if (string.IsNullOrWhiteSpace(ownerPassword)) {
                    throw new InvalidOperationException("The current owner password is required to replace this document's protection.");
                }
                return document.Security.Reencrypt(ownerPassword, encryption).Pdf;
            }, cancellationToken).ConfigureAwait(false);
            progress?.Report(new PdfWorkspaceProgress("Writing protected copy", 0.8D));
            await WriteOutputAsync(destination, output, cancellationToken).ConfigureAwait(false);
            progress?.Report(new PdfWorkspaceProgress("Protected copy saved", 1D));
        } finally {
            _operationGate.Release();
        }
    }

    internal async Task SaveDecryptedCopyAsync(
        string destinationPath,
        string? ownerPassword,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (!IsEncrypted) throw new InvalidOperationException("This PDF is not password protected.");
        if (!CanChangeEncryption(ownerPassword)) {
            throw new InvalidOperationException("This document's signature, certification, usage-rights, or authorization policy prevents removing password protection.");
        }
        string destination = ValidateExportDestination(destinationPath);
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            string effectiveOwnerPassword = _documentInfo.Security.HasOwnerAuthorization
                ? _readOptions.Password ?? string.Empty
                : ownerPassword ?? string.Empty;
            if (string.IsNullOrWhiteSpace(effectiveOwnerPassword)) {
                throw new InvalidOperationException("The owner password is required to remove document protection.");
            }
            progress?.Report(new PdfWorkspaceProgress("Removing password protection", 0.15D));
            byte[] output = await RunCancellableCpuWorkAsync(
                () => LoadDocument(_bytes).Security.Decrypt(effectiveOwnerPassword).Pdf,
                cancellationToken).ConfigureAwait(false);
            progress?.Report(new PdfWorkspaceProgress("Writing decrypted copy", 0.8D));
            await WriteOutputAsync(destination, output, cancellationToken).ConfigureAwait(false);
            progress?.Report(new PdfWorkspaceProgress("Decrypted copy saved", 1D));
        } finally {
            _operationGate.Release();
        }
    }

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
        string destination = System.IO.Path.GetFullPath(destinationPath);
        StringComparison comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        if (string.Equals(destination, Path, comparison)) {
            throw new InvalidOperationException("Choose a different output path so the open document remains unchanged.");
        }
        return destination;
    }

    private static async Task WriteOutputAsync(string destination, byte[] bytes, CancellationToken cancellationToken) {
        string? directory = System.IO.Path.GetDirectoryName(destination);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory)) {
            throw new DirectoryNotFoundException("The output folder does not exist.");
        }
        string temporaryPath = System.IO.Path.Combine(
            directory,
            "." + System.IO.Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try {
            await File.WriteAllBytesAsync(temporaryPath, bytes, cancellationToken).ConfigureAwait(false);
            File.Move(temporaryPath, destination, overwrite: true);
        } finally {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }
}
