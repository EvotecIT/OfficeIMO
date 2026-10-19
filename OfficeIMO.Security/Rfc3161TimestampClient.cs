using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Org.BouncyCastle.Math;
using Org.BouncyCastle.Tsp;

namespace OfficeIMO.Security;

/// <summary>RFC 3161 client with caller-owned HTTP transport, endpoint, credentials and trust policy.</summary>
public sealed class Rfc3161TimestampClient {
    private readonly Func<byte[], CancellationToken, Task<byte[]>> _transport;
    private readonly CertificateValidationOptions _validation;
    private readonly int _maximumResponseBytes;

    /// <summary>
    /// Creates a client. The transport must POST the request as application/timestamp-query and return
    /// the application/timestamp-reply body, enforcing an HTTP response-size limit before buffering it.
    /// The client checks the response again and verifies nonce, imprint, signature, TSA profile and trust.
    /// </summary>
    public Rfc3161TimestampClient(Func<byte[], CancellationToken, Task<byte[]>> transport,
        CertificateValidationOptions? certificateValidation = null, int maximumResponseBytes = 1024 * 1024) {
        _transport = transport ?? throw new ArgumentNullException(nameof(transport));
        _validation = certificateValidation ?? new CertificateValidationOptions();
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumResponseBytes);
#else
        if (maximumResponseBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumResponseBytes));
#endif
        _maximumResponseBytes = maximumResponseBytes;
    }

    /// <summary>Requests a SHA-256 timestamp over the supplied bytes, returning a verified DER token.</summary>
    public async Task<byte[]> GetTimestampAsync(byte[] timestampedData, CancellationToken cancellationToken = default) {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(timestampedData);
#else
        if (timestampedData is null) throw new ArgumentNullException(nameof(timestampedData));
#endif
        cancellationToken.ThrowIfCancellationRequested();
#if NET8_0_OR_GREATER
        byte[] digest = SHA256.HashData(timestampedData);
#else
        byte[] digest;
        using (SHA256 sha256 = SHA256.Create()) digest = sha256.ComputeHash(timestampedData);
#endif
        byte[] nonce = new byte[20];
        using (RandomNumberGenerator random = RandomNumberGenerator.Create()) random.GetBytes(nonce);
        var generator = new TimeStampRequestGenerator();
        generator.SetCertReq(true);
        TimeStampRequest request = generator.Generate("2.16.840.1.101.3.4.2.1", digest, new BigInteger(1, nonce));
        byte[] responseBytes = await _transport(request.GetEncoded(), cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (responseBytes is null) throw new CryptographicException("The timestamp transport returned no response.");
        SecurityLimits.EnsureBufferWithinLimit(responseBytes, _maximumResponseBytes, nameof(responseBytes));
        var response = new TimeStampResponse(responseBytes);
        response.Validate(request);
        TimeStampToken token = response.TimeStampToken ?? throw new CryptographicException("The TSA did not grant a timestamp.");
        byte[] encoded = token.GetEncoded();
        Rfc3161TimestampVerificationResult result = Rfc3161TimestampVerifier.Verify(encoded, timestampedData, _validation,
            _maximumResponseBytes);
        if (result.Status != SecurityValidationStatus.Valid) {
            throw new CryptographicException("TSA token validation failed: " + string.Join("; ", result.Findings.Select(finding => finding.Message)));
        }
        return encoded;
    }

    /// <summary>Synchronous adapter for CMS signing's timestamp callback.</summary>
    public byte[] GetTimestamp(byte[] timestampedData) =>
        // The transport may capture its invocation context before its first await.
        // Start it on the thread pool so a synchronous UI signer cannot block that context.
        Task.Run(() => GetTimestampAsync(timestampedData)).GetAwaiter().GetResult();
}
