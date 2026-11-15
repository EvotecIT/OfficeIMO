using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
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
    public bool AllowNonPublicAddresses { get; set; }
}

/// <summary>
/// Fetches OCSP responses (RFC 6960) and CRLs (RFC 5280) from the access locations named in a certificate's
/// authority information access and CRL distribution points extensions.
/// </summary>
/// <remarks>
/// OCSP is tried first; CRLs are fetched when no OCSP location returns a status signed by the issuer or its
/// responder. Responses are cached by the instance until their nextUpdate. Only http and https locations are used
/// and connections go directly to the server, without a proxy. Non-public addresses are refused unless
/// <see cref="CertificateRevocationFetcherOptions.AllowNonPublicAddresses"/> is set; on .NET 8 and later the
/// address is checked when the connection is made, on older frameworks the request is sent to the checked address
/// and https locations are refused.
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
            AllowNonPublicAddresses = options.AllowNonPublicAddresses
        };
        _client = new HttpClient(CreateHandler(_options.AllowNonPublicAddresses), disposeHandler: true) {
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
                IPAddress? pinned = await EnsureAllowedAsync(current, timeout.Token).ConfigureAwait(false);
                using HttpRequestMessage request = CreateRequest(current, pinned, ocspRequest);
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

    private static HttpRequestMessage CreateRequest(Uri location, IPAddress? pinned, byte[]? ocspRequest) {
        Uri target = location;
        string? host = null;
        if (pinned is not null) {
            // Connect to the checked address and name the original host, so a second name lookup cannot
            // redirect the request to an internal address.
            target = new UriBuilder(location) {
                Host = pinned.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6 ? "[" + pinned + "]" : pinned.ToString()
            }.Uri;
            host = location.IsDefaultPort ? location.Host : location.Host + ":" + location.Port.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        var request = new HttpRequestMessage(ocspRequest is null ? HttpMethod.Get : HttpMethod.Post, target);
        if (host is not null) request.Headers.Host = host;
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

    /// <summary>
    /// Checks the location and, on frameworks without a connect callback, returns the checked public address
    /// the request must be sent to. Returns null when the handler checks addresses itself or checks are off.
    /// </summary>
    private async Task<IPAddress?> EnsureAllowedAsync(Uri location, CancellationToken cancellationToken) {
        if (!location.IsAbsoluteUri || (location.Scheme != Uri.UriSchemeHttp && location.Scheme != Uri.UriSchemeHttps)) {
            throw new HttpRequestException("Only http and https locations are fetched.");
        }
        if (!string.IsNullOrEmpty(location.UserInfo)) {
            throw new HttpRequestException("Locations with user information are not fetched.");
        }
        if (_options.AllowNonPublicAddresses) return null;
        IPAddress[] addresses = IPAddress.TryParse(location.DnsSafeHost, out IPAddress? literal)
            ? new[] { literal }
#if NET6_0_OR_GREATER
            : await Dns.GetHostAddressesAsync(location.DnsSafeHost, cancellationToken).ConfigureAwait(false);
#else
            : await Dns.GetHostAddressesAsync(location.DnsSafeHost).ConfigureAwait(false);
#endif
        if (addresses.Length == 0) throw new HttpRequestException("The host name did not resolve.");
        foreach (IPAddress address in addresses) {
            if (!PublicNetworkAddress.IsPublic(address)) {
                throw new HttpRequestException("The location resolves to the non-public address " + address + ".");
            }
        }
#if NET8_0_OR_GREATER
        return null;
#else
        if (location.Scheme == Uri.UriSchemeHttps) {
            throw new HttpRequestException("On this framework an https location cannot be bound to its checked public address, so it is not fetched.");
        }
        return addresses[0];
#endif
    }

#if NET8_0_OR_GREATER
    private static SocketsHttpHandler CreateHandler(bool allowNonPublicAddresses) {
        var handler = new SocketsHttpHandler {
            AllowAutoRedirect = false,
            UseProxy = false,
            UseCookies = false,
            AutomaticDecompression = DecompressionMethods.None
        };
        if (!allowNonPublicAddresses) {
            handler.ConnectCallback = async (context, cancellationToken) => {
                IPAddress[] addresses = IPAddress.TryParse(context.DnsEndPoint.Host, out IPAddress? literal)
                    ? new[] { literal }
                    : await Dns.GetHostAddressesAsync(context.DnsEndPoint.Host, cancellationToken).ConfigureAwait(false);
                if (addresses.Length == 0 || addresses.Any(candidate => !PublicNetworkAddress.IsPublic(candidate))) {
                    throw new HttpRequestException("The location resolves to a non-public address.");
                }
                Exception? last = null;
                foreach (IPAddress address in addresses) {
                    var socket = new System.Net.Sockets.Socket(address.AddressFamily, System.Net.Sockets.SocketType.Stream,
                        System.Net.Sockets.ProtocolType.Tcp) { NoDelay = true };
                    try {
                        await socket.ConnectAsync(new IPEndPoint(address, context.DnsEndPoint.Port), cancellationToken).ConfigureAwait(false);
                        return new System.Net.Sockets.NetworkStream(socket, ownsSocket: true);
                    } catch (System.Net.Sockets.SocketException exception) {
                        socket.Dispose();
                        last = exception;
                    } catch {
                        socket.Dispose();
                        throw;
                    }
                }
                throw new HttpRequestException("No checked address of the location accepted a connection.", last);
            };
        }
        return handler;
    }
#else
    private static HttpClientHandler CreateHandler(bool allowNonPublicAddresses) {
        // Older frameworks have no connect callback; requests are sent to the address EnsureAllowedAsync checked.
        _ = allowNonPublicAddresses;
        return new HttpClientHandler {
            AllowAutoRedirect = false,
            UseProxy = false,
            UseCookies = false,
            AutomaticDecompression = DecompressionMethods.None
        };
    }
#endif
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

/// <summary>Classifies addresses that a revocation fetcher may contact without explicit permission.</summary>
internal static class PublicNetworkAddress {
    internal static bool IsPublic(IPAddress address) {
        if (address.IsIPv4MappedToIPv6) return IsPublic(address.MapToIPv4());
        byte[] bytes = address.GetAddressBytes();
        if (address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork) return IsPublicIPv4(bytes);
        if (address.AddressFamily != System.Net.Sockets.AddressFamily.InterNetworkV6 || bytes.Length != 16) return false;
        if (IPAddress.IsLoopback(address) || address.Equals(IPAddress.IPv6None) || address.Equals(IPAddress.IPv6Any)) return false;
        if (address.IsIPv6LinkLocal || address.IsIPv6SiteLocal || address.IsIPv6Multicast) return false;
        if ((bytes[0] & 0xFE) == 0xFC) return false;                                      // fc00::/7 unique local
        if (bytes[0] == 0x20 && bytes[1] == 0x01 && bytes[2] == 0x0D && bytes[3] == 0xB8) return false; // 2001:db8::/32 documentation
        if (bytes[0] == 0x20 && bytes[1] == 0x01 && bytes[2] < 0x02) return false;       // 2001::/23 protocol assignments, Teredo
        if (bytes[0] == 0x01 && bytes[1] == 0x00 && bytes.Skip(2).Take(6).All(static b => b == 0)) return false; // 100::/64 discard
        if (bytes[0] == 0x20 && bytes[1] == 0x02) {                                       // 2002::/16 6to4 embeds IPv4
            return IsPublicIPv4(new[] { bytes[2], bytes[3], bytes[4], bytes[5] });
        }
        if (bytes[0] == 0x00 && bytes[1] == 0x64 && bytes[2] == 0xFF && bytes[3] == 0x9B && bytes[4] == 0x00 && bytes[5] == 0x01) {
            return false;                                                                  // 64:ff9b:1::/48 local-use translation
        }
        if (bytes[0] == 0x00 && bytes[1] == 0x64 && bytes[2] == 0xFF && bytes[3] == 0x9B &&
            bytes.Skip(4).Take(8).All(static b => b == 0)) {                              // 64:ff9b::/96 NAT64 embeds IPv4
            return IsPublicIPv4(new[] { bytes[12], bytes[13], bytes[14], bytes[15] });
        }
        return true;
    }

    private static bool IsPublicIPv4(byte[] b) {
        if (b.Length != 4) return false;
        return !(b[0] == 0 ||                                   // 0.0.0.0/8
            b[0] == 10 ||                                       // 10.0.0.0/8
            (b[0] == 100 && (b[1] & 0xC0) == 64) ||             // 100.64.0.0/10 shared address space
            b[0] == 127 ||                                      // loopback
            (b[0] == 169 && b[1] == 254) ||                     // link-local
            (b[0] == 172 && (b[1] & 0xF0) == 16) ||             // 172.16.0.0/12
            (b[0] == 192 && b[1] == 0 && b[2] == 0) ||          // 192.0.0.0/24 protocol assignments
            (b[0] == 192 && b[1] == 0 && b[2] == 2) ||          // TEST-NET-1
            (b[0] == 192 && b[1] == 168) ||                     // 192.168.0.0/16
            (b[0] == 198 && (b[1] & 0xFE) == 18) ||             // 198.18.0.0/15 benchmarking
            (b[0] == 198 && b[1] == 51 && b[2] == 100) ||       // TEST-NET-2
            (b[0] == 203 && b[1] == 0 && b[2] == 113) ||        // TEST-NET-3
            b[0] >= 224);                                       // multicast, reserved, broadcast
    }
}
