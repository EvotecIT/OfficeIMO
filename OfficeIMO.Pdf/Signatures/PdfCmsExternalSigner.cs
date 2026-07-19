using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>First-party PDF detached-CMS signer backed by the shared OfficeIMO security engine.</summary>
public sealed class PdfCmsExternalSigner : IPdfExternalSigner, IDisposable {
    private readonly X509Certificate2 _certificate;
    private readonly List<X509Certificate2> _certificateChain;
    private bool _disposed;

    /// <summary>Creates an RSA CMS signer and clones the certificate handles used by it.</summary>
    public PdfCmsExternalSigner(
        X509Certificate2 certificate,
        string? name = null,
        CmsSigningOptions? signingOptions = null,
        IEnumerable<X509Certificate2>? certificateChain = null) {
#if NETSTANDARD2_0 || NET472
        if (certificate == null) throw new ArgumentNullException(nameof(certificate));
#else
        ArgumentNullException.ThrowIfNull(certificate);
#endif
        if (!certificate.HasPrivateKey) {
            throw new ArgumentException("The signing certificate must include a private key.", nameof(certificate));
        }
        _certificate = new X509Certificate2(certificate);
        _certificateChain = certificateChain?.Select(static item => new X509Certificate2(item)).ToList()
            ?? new List<X509Certificate2>();
        SigningOptions = signingOptions ?? new CmsSigningOptions();
        Name = string.IsNullOrWhiteSpace(name) ? "OfficeIMO.Pdf CMS signer" : name!;
    }

    /// <inheritdoc />
    public string Name { get; }

    /// <summary>Shared CMS signing options used for every signature produced by this instance.</summary>
    public CmsSigningOptions SigningOptions { get; }

    /// <inheritdoc />
    public byte[] Sign(PdfExternalSignatureRequest request) {
#if NETSTANDARD2_0 || NET472
        if (request == null) throw new ArgumentNullException(nameof(request));
#else
        ArgumentNullException.ThrowIfNull(request);
#endif
        ThrowIfDisposed();
        return CmsSignedDataSigner.SignDetached(
            request.SignedContent,
            _certificate,
            SigningOptions,
            _certificateChain);
    }

    /// <summary>Releases cloned certificate and private-key handles.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _certificate.Dispose();
        foreach (X509Certificate2 certificate in _certificateChain) certificate.Dispose();
    }

    private void ThrowIfDisposed() {
#if NETSTANDARD2_0 || NET472
        if (_disposed) throw new ObjectDisposedException(nameof(PdfCmsExternalSigner));
#else
        ObjectDisposedException.ThrowIf(_disposed, this);
#endif
    }
}
