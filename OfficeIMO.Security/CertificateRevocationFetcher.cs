using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Core.Internal;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Ocsp;
using BcX509Certificate = Org.BouncyCastle.X509.X509Certificate;

namespace OfficeIMO.Security;

/// <summary>Limits and network policy for <see cref="CertificateRevocationFetcher"/>.</summary>
public sealed class CertificateRevocationFetcherOptions {
    /// <summary>Time allowed for one HTTP exchange, including redirects. Defaults to 15 seconds.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(15);

    /// <summary>Largest accepted OCSP response or CRL body. Defaults to 32 MiB.</summary>
    public long MaxResponseBytes { get; set; } = 32L * 1024 * 1024;

    /// <summary>
    /// Allows loopback, private, link-local and other non-public addresses. Defaults to false, because the
    /// addresses come from certificates inside the document being processed. Enable it only for an internal PKI.
    /// </summary>
    public bool AllowPrivateNetworkAddresses { get; set; }
}

/// <summary>
/// Fetches OCSP responses (RFC 6960) and CRLs (RFC 5280) from the access locations named in a certificate's
/// authority information access and CRL distribution points extensions.
/// </summary>
/// <remarks>
/// OCSP is tried first; CRLs are fetched when no OCSP location returns a status signed by the issuer or its
/// responder. Responses are cached by the instance until their nextUpdate. Only http and https locations are used
/// and, while non-public addresses are refused, connections go directly to the server without a proxy. Non-public
/// addresses are refused unless <see cref="CertificateRevocationFetcherOptions.AllowPrivateNetworkAddresses"/> is
/// set; on .NET 8 and later every address is checked when the connection is made. Older frameworks check the
/// resolved addresses before each request, send http requests to the checked address and rely on certificate
/// validation to bind https connections to the host name.
/// </remarks>
public sealed class CertificateRevocationFetcher : ICertificateRevocationFetcher, IDisposable {
    private const string OcspAccessMethodOid = "1.3.6.1.5.5.7.48.1";
    private const int MaxRedirects = 3;
    private static readonly TimeSpan UndatedResponseLifetime = TimeSpan.FromMinutes(5);
    private static readonly TimeSpan FailureLifetime = TimeSpan.FromMinutes(1);
    private const int MaxCacheEntries = 256;
    private readonly CertificateRevocationFetcherOptions _options;
    private readonly HttpClient _client;
    private readonly CancellationTokenSource _disposal = new();
    private readonly Dictionary<string, CacheEntry> _cache = new(StringComparer.Ordinal);
    private readonly object _cacheLock = new();
    private long _cacheSequence;
    private bool _disposed;

    /// <summary>Creates a fetcher. The options are copied.</summary>
    public CertificateRevocationFetcher(CertificateRevocationFetcherOptions? options = null) {
        options ??= new CertificateRevocationFetcherOptions();
        if (options.Timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(options), "Timeout must be positive.");
        if (options.MaxResponseBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxResponseBytes must be positive.");
        _options = new CertificateRevocationFetcherOptions {
            Timeout = options.Timeout,
            MaxResponseBytes = options.MaxResponseBytes,
            AllowPrivateNetworkAddresses = options.AllowPrivateNetworkAddresses
        };
        _client = new HttpClient(
            OfficeNetworkAddressPolicy.CreateHandler(_options.AllowPrivateNetworkAddresses, DecompressionMethods.None),
            disposeHandler: true) {
            Timeout = System.Threading.Timeout.InfiniteTimeSpan
        };
    }

    /// <inheritdoc />
    public async Task<CertificateRevocationEvidence> FetchAsync(
        X509Certificate2 certificate,
        X509Certificate2 issuer,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(certificate);
        ArgumentNullException.ThrowIfNull(issuer);
        ThrowIfDisposed();
        cancellationToken.ThrowIfCancellationRequested();

        var parser = new Org.BouncyCastle.X509.X509CertificateParser();
        BcX509Certificate subject = parser.ReadCertificate(certificate.RawData);
        BcX509Certificate issuerCertificate = parser.ReadCertificate(issuer.RawData);
        string label = "'" + subject.SubjectDN + "'";
        var findings = new List<SecurityFinding>();
        var ocspResponses = new List<byte[]>();
        var crls = new List<byte[]>();

        List<Uri> ocspLocations = ReadOcspLocations(subject, findings, label);
        List<Uri> crlLocations = ReadCrlLocations(subject, findings, label);
        if (ocspLocations.Count == 0 && crlLocations.Count == 0) {
            findings.Add(Finding("RevocationLocationMissing",
                "Certificate " + label + " names no OCSP responder or CRL distribution point."));
        }

        // Stop at the first source whose answer is signed by the issuer or its responder; an unusable answer
        // must not hide the remaining responders and distribution points.
        bool settled = false;
        if (ocspLocations.Count > 0) {
            byte[] request = CreateOcspRequest(subject, issuerCertificate);
            foreach (Uri location in ocspLocations) {
                byte[]? response = await TryGetAsync(location, request, label, findings, cancellationToken).ConfigureAwait(false);
                if (response is null) continue;
                if (!HasVerifiedStatus(response, subject, issuerCertificate, out string? problem)) {
                    findings.Add(Finding("RevocationOcspNoStatus",
                        "The OCSP responder " + location + " returned no usable status for " + label + ": " + problem));
                    continue;
                }
                ocspResponses.Add(response);
                settled = true;
                break;
            }
        }

        if (!settled) {
            foreach (Uri location in crlLocations) {
                byte[]? response = await TryGetAsync(location, null, label, findings, cancellationToken).ConfigureAwait(false);
                if (response is null) continue;
                Org.BouncyCastle.X509.X509Crl? crl;
                try {
                    crl = new Org.BouncyCastle.X509.X509CrlParser().ReadCrl(response) ?? throw new InvalidDataException("No CRL was decoded.");
                } catch (Exception exception) when (IsRecoverable(exception)) {
                    findings.Add(Finding("RevocationTokenMalformed",
                        "The CRL from " + location + " could not be decoded: " + exception.Message));
                    continue;
                }
                crls.Add(response);
                if (IsCurrentCrlFrom(crl, issuerCertificate)) break;
                findings.Add(Finding("RevocationCrlNotCurrent",
                    "The CRL from " + location + " is not signed by the issuer of " + label + " or is past its next update."));
            }
        }

        return new CertificateRevocationEvidence(new[] { issuer.RawData }, ocspResponses, crls, findings);
    }

    /// <summary>Releases the HTTP connection pool and abandons requests still in flight.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _disposal.Cancel();
        _client.Dispose();
        _disposal.Dispose();
    }

    private async Task<byte[]?> TryGetAsync(
        Uri location,
        byte[]? ocspRequest,
        string label,
        List<SecurityFinding> findings,
        CancellationToken cancellationToken) {
        string key = (ocspRequest is null ? "GET " : "POST " + Convert.ToBase64String(ocspRequest) + " ") + location.AbsoluteUri;
        CacheEntry entry;
        lock (_cacheLock) {
            // A cached answer is reused until its nextUpdate, so a long-lived fetcher does not serve stale data.
            DateTime now = DateTime.UtcNow;
            if (!_cache.TryGetValue(key, out entry!) || entry.IsExpired(now)) {
                TrimCache(now);
                entry = new CacheEntry { Sequence = ++_cacheSequence };
                entry.Task = FetchAndStampAsync(entry, location, ocspRequest);
                _cache[key] = entry;
            }
        }
        try {
            // Each caller may stop waiting; the shared request keeps running for the others.
            return await WaitAsync(entry.Task, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) when (IsRecoverable(exception)) {
            findings.Add(Finding("RevocationFetchFailed",
                "Revocation data for " + label + " could not be fetched from " + location + ": " + exception.Message));
            return null;
        }
    }

    /// <summary>Drops expired answers and, above the entry limit, the oldest completed ones; requests in flight stay.</summary>
    private void TrimCache(DateTime now) {
        foreach (string expired in _cache.Where(item => item.Value.IsExpired(now)).Select(static item => item.Key).ToList()) {
            _cache.Remove(expired);
        }
        if (_cache.Count < MaxCacheEntries) return;
        foreach (string oldest in _cache.Where(static item => item.Value.Task.IsCompleted)
                     .OrderBy(static item => item.Value.Sequence)
                     .Take(_cache.Count - MaxCacheEntries + 1)
                     .Select(static item => item.Key)
                     .ToList()) {
            _cache.Remove(oldest);
        }
    }

    private async Task<byte[]> FetchAndStampAsync(CacheEntry entry, Uri location, byte[]? ocspRequest) {
        try {
            byte[] response = await SendAsync(location, ocspRequest).ConfigureAwait(false);
            entry.ExpiresUtc = ReadNextUpdate(response, ocspRequest is not null) ?? DateTime.UtcNow + UndatedResponseLifetime;
            return response;
        } catch {
            entry.ExpiresUtc = DateTime.UtcNow + FailureLifetime;
            throw;
        }
    }

    private static async Task<T> WaitAsync<T>(Task<T> task, CancellationToken cancellationToken) {
#if NET6_0_OR_GREATER
        return await task.WaitAsync(cancellationToken).ConfigureAwait(false);
#else
        if (!cancellationToken.CanBeCanceled) return await task.ConfigureAwait(false);
        var cancelled = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using (cancellationToken.Register(static state => ((TaskCompletionSource<bool>)state!).TrySetResult(true), cancelled)) {
            if (await Task.WhenAny(task, cancelled.Task).ConfigureAwait(false) != task) {
                cancellationToken.ThrowIfCancellationRequested();
            }
        }
        return await task.ConfigureAwait(false);
#endif
    }

    private async Task<byte[]> SendAsync(Uri location, byte[]? ocspRequest) {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(_disposal.Token);
        timeout.CancelAfter(_options.Timeout);
        try {
            Uri current = location;
            for (int redirect = 0; ; redirect++) {
                using HttpRequestMessage request = await CreateRequestAsync(current, ocspRequest, timeout.Token).ConfigureAwait(false);
                using HttpResponseMessage response = await _client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token)
                    .ConfigureAwait(false);
                int code = (int)response.StatusCode;
                if (code is 301 or 302 or 303 or 307 or 308) {
                    Uri? target = response.Headers.Location;
                    if (target is null) throw new HttpRequestException("The server redirected without a location.");
                    if (redirect >= MaxRedirects) throw new HttpRequestException("The server redirected more than " + MaxRedirects + " times.");
                    current = target.IsAbsoluteUri ? target : new Uri(current, target);
                    if (ocspRequest is not null && code == 303) throw new HttpRequestException("The OCSP responder redirected a POST to a GET.");
                    continue;
                }
                if (!response.IsSuccessStatusCode) {
                    throw new HttpRequestException("The server returned HTTP " + code + ".");
                }
                return await ReadBoundedAsync(response.Content, timeout.Token).ConfigureAwait(false);
            }
        } catch (OperationCanceledException) when (!_disposal.IsCancellationRequested) {
            throw new TimeoutException("The request did not complete within " + _options.Timeout.TotalSeconds.ToString(System.Globalization.CultureInfo.InvariantCulture) + " seconds.");
        }
    }

    private async Task<HttpRequestMessage> CreateRequestAsync(Uri location, byte[]? ocspRequest, CancellationToken cancellationToken) {
        if (!location.IsAbsoluteUri || (location.Scheme != Uri.UriSchemeHttp && location.Scheme != Uri.UriSchemeHttps)) {
            throw new HttpRequestException("Only http and https locations are fetched.");
        }
        if (!string.IsNullOrEmpty(location.UserInfo)) {
            throw new HttpRequestException("Locations with user information are not fetched.");
        }
        var request = await OfficeNetworkAddressPolicy.CreateRequestAsync(
            ocspRequest is null ? HttpMethod.Get : HttpMethod.Post, location,
            _options.AllowPrivateNetworkAddresses, cancellationToken).ConfigureAwait(false);
        if (ocspRequest is not null) {
            request.Content = new ByteArrayContent(ocspRequest);
            request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/ocsp-request");
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/ocsp-response"));
        }
        return request;
    }

    private async Task<byte[]> ReadBoundedAsync(HttpContent content, CancellationToken cancellationToken) {
        long limit = _options.MaxResponseBytes;
        if (content.Headers.ContentLength is long declared && declared > limit) {
            throw new InvalidDataException("The response declares " + declared + " bytes, more than the limit of " + limit + " bytes.");
        }
#if NET5_0_OR_GREATER
        using Stream stream = await content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
#else
        using Stream stream = await content.ReadAsStreamAsync().ConfigureAwait(false);
#endif
        using var buffer = new MemoryStream();
        byte[] chunk = new byte[81920];
        while (true) {
#if NET5_0_OR_GREATER
            int read = await stream.ReadAsync(chunk.AsMemory(0, chunk.Length), cancellationToken).ConfigureAwait(false);
#else
            int read = await stream.ReadAsync(chunk, 0, chunk.Length, cancellationToken).ConfigureAwait(false);
#endif
            if (read == 0) break;
            if (buffer.Length + read > limit) {
                throw new InvalidDataException("The response is larger than the limit of " + limit + " bytes.");
            }
            buffer.Write(chunk, 0, read);
        }
        if (buffer.Length == 0) throw new InvalidDataException("The response is empty.");
        return buffer.ToArray();
    }

    private static byte[] CreateOcspRequest(BcX509Certificate certificate, BcX509Certificate issuer) {
        // SHA-1 CertID is the interoperable choice for lightweight responders (RFC 5019 section 2.1.1).
        var generator = new OcspReqGenerator();
        generator.AddRequest(new CertificateID(CertificateID.DigestSha1, issuer, certificate.SerialNumber));
        return generator.Generate().GetEncoded();
    }

    private static bool HasVerifiedStatus(byte[] encoded, BcX509Certificate certificate, BcX509Certificate issuer, out string? problem) {
        try {
            var response = new OcspResp(encoded);
            if (response.Status != OcspRespStatus.Successful) {
                problem = "response status " + response.Status + ".";
                return false;
            }
            if (response.GetResponseObject() is not BasicOcspResp basic) {
                problem = "not a basic OCSP response.";
                return false;
            }
            if (!CertificateRevocationEvaluator.HasVerifiedStatus(basic, certificate, issuer)) {
                problem = "the response does not cover the certificate or is not signed by its issuer or an authorized responder.";
                return false;
            }
            problem = null;
            return true;
        } catch (Exception exception) when (IsRecoverable(exception)) {
            problem = exception.Message;
            return false;
        }
    }

    private static bool IsCurrentCrlFrom(Org.BouncyCastle.X509.X509Crl crl, BcX509Certificate issuer) {
        if (!crl.IssuerDN.Equivalent(issuer.SubjectDN, inOrder: true)) return false;
        try {
            crl.Verify(issuer.GetPublicKey());
        } catch (Exception exception) when (IsRecoverable(exception)) {
            return false;
        }
        return crl.NextUpdate is not DateTime next || next.ToUniversalTime() >= DateTime.UtcNow;
    }

    private static DateTime? ReadNextUpdate(byte[] encoded, bool isOcsp) {
        try {
            if (isOcsp) {
                if (new OcspResp(encoded).GetResponseObject() is not BasicOcspResp basic) return null;
                DateTime? earliest = null;
                foreach (SingleResp single in basic.Responses) {
                    if (single.NextUpdate is not DateTime next) return null;
                    next = next.ToUniversalTime();
                    if (!earliest.HasValue || next < earliest.Value) earliest = next;
                }
                return earliest;
            }
            return new Org.BouncyCastle.X509.X509CrlParser().ReadCrl(encoded)?.NextUpdate?.ToUniversalTime();
        } catch (Exception exception) when (IsRecoverable(exception)) {
            return null;
        }
    }

    private sealed class CacheEntry {
        internal Task<byte[]> Task { get; set; } = null!;
        internal DateTime? ExpiresUtc { get; set; }
        internal long Sequence { get; set; }

        internal bool IsExpired(DateTime now) => ExpiresUtc is DateTime expires && now >= expires;
    }
    private static List<Uri> ReadOcspLocations(BcX509Certificate certificate, List<SecurityFinding> findings, string label) {
        var result = new List<Uri>();
        try {
            Asn1OctetString? value = certificate.GetExtensionValue(X509Extensions.AuthorityInfoAccess);
            if (value is null) return result;
            AuthorityInformationAccess access = AuthorityInformationAccess.GetInstance(
                Asn1Object.FromByteArray(value.GetOctets()));
            foreach (AccessDescription description in access.GetAccessDescriptions()) {
                if (description.AccessMethod.Id != OcspAccessMethodOid) continue;
                AddLocation(description.AccessLocation, result);
            }
        } catch (Exception exception) when (IsRecoverable(exception)) {
            findings.Add(Finding("RevocationLocationMalformed",
                "The authority information access extension of " + label + " could not be read: " + exception.Message));
        }
        return result;
    }

    private static List<Uri> ReadCrlLocations(BcX509Certificate certificate, List<SecurityFinding> findings, string label) {
        var result = new List<Uri>();
        try {
            Asn1OctetString? value = certificate.GetExtensionValue(X509Extensions.CrlDistributionPoints);
            if (value is null) return result;
            CrlDistPoint points = CrlDistPoint.GetInstance(Asn1Object.FromByteArray(value.GetOctets()));
            foreach (DistributionPoint point in points.GetDistributionPoints()) {
                // Indirect CRLs (with a cRLIssuer) are not fetched; the validator checks direct CRLs only.
                if (point.CrlIssuer is not null) continue;
                if (point.DistributionPointName?.Type != DistributionPointName.FullName) continue;
                foreach (GeneralName name in GeneralNames.GetInstance(point.DistributionPointName.Name).GetNames()) {
                    AddLocation(name, result);
                }
            }
        } catch (Exception exception) when (IsRecoverable(exception)) {
            findings.Add(Finding("RevocationLocationMalformed",
                "The CRL distribution points extension of " + label + " could not be read: " + exception.Message));
        }
        return result;
    }

    private static void AddLocation(GeneralName name, List<Uri> result) {
        if (name.TagNo != GeneralName.UniformResourceIdentifier) return;
        string text = DerIA5String.GetInstance(name.Name).GetString();
        if (Uri.TryCreate(text, UriKind.Absolute, out Uri? location) &&
            (location.Scheme == Uri.UriSchemeHttp || location.Scheme == Uri.UriSchemeHttps) &&
            !result.Contains(location)) {
            result.Add(location);
        }
    }

    private void ThrowIfDisposed() {
#if NET8_0_OR_GREATER
        ObjectDisposedException.ThrowIf(_disposed, this);
#else
        if (_disposed) throw new ObjectDisposedException(nameof(CertificateRevocationFetcher));
#endif
    }

    private static SecurityFinding Finding(string code, string message) =>
        new SecurityFinding(SecurityFindingSeverity.Warning, code, message);

    private static bool IsRecoverable(Exception exception) =>
        exception is not OutOfMemoryException &&
        exception is not StackOverflowException &&
        exception is not AccessViolationException &&
        exception is not OperationCanceledException;
}
