using System.Buffers.Binary;
using System.Collections.ObjectModel;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Serialization;

namespace OfficeIMO.Html.Runtime;

/// <summary>A normalized dynamic HTTP request that can be acquired once and replayed without worker network access.</summary>
public sealed class HtmlRuntimeFetchRequest {
    private readonly byte[]? _body;

    /// <summary>Creates an immutable request identity from browser-visible fetch options.</summary>
    public HtmlRuntimeFetchRequest(Uri url, string method = "GET", IReadOnlyDictionary<string, string>? headers = null,
        byte[]? body = null, string mode = "cors", string credentials = "same-origin", string redirect = "follow")
        : this(url, method, headers, body, mode, credentials, redirect, body != null) { }

    /// <summary>Restores an immutable request identity while preserving the distinction between no body and a zero-byte body.</summary>
    [JsonConstructor]
    public HtmlRuntimeFetchRequest(Uri url, string method, IReadOnlyDictionary<string, string>? headers,
        byte[]? body, string mode, string credentials, string redirect, bool hasBody) {
        Url = HtmlRuntimeResourcePolicy.ValidateUrl(url);
        Method = NormalizeMethod(method);
        HasBody = hasBody;
        if (HasBody && Method is "GET" or "HEAD") throw new ArgumentException("GET and HEAD requests cannot have a body.", nameof(body));
        if (mode is not ("cors" or "same-origin")) throw new ArgumentException("Unsupported fetch mode.", nameof(mode));
        if (credentials is not ("omit" or "same-origin")) throw new ArgumentException("Unsupported fetch credentials mode.", nameof(credentials));
        if (redirect is not ("follow" or "error")) throw new ArgumentException("Unsupported fetch redirect mode.", nameof(redirect));
        Headers = new ReadOnlyDictionary<string, string>(NormalizeHeaders(headers));
        _body = HasBody ? (byte[])(body ?? Array.Empty<byte>()).Clone() : null;
        Mode = mode;
        Credentials = credentials;
        Redirect = redirect;
        Identity = ComputeIdentity();
    }

    /// <summary>Absolute request URL.</summary>
    public Uri Url { get; }
    /// <summary>Uppercase HTTP method.</summary>
    public string Method { get; }
    /// <summary>Normalized outgoing request headers after browser-forbidden names were removed.</summary>
    public IReadOnlyDictionary<string, string> Headers { get; }
    /// <summary>Independent request body bytes, or <see langword="null"/>.</summary>
    public byte[]? Body => _body == null ? null : (byte[])_body.Clone();
    /// <summary>Whether the request has a body, including a zero-byte body.</summary>
    public bool HasBody { get; }
    /// <summary>Fetch request mode.</summary>
    public string Mode { get; }
    /// <summary>Fetch credentials mode. The managed transport never owns a cookie jar.</summary>
    public string Credentials { get; }
    /// <summary>Fetch redirect mode.</summary>
    public string Redirect { get; }
    /// <summary>Stable lowercase SHA-256 identity over URL, method, headers, body, and fetch options.</summary>
    public string Identity { get; }
    /// <summary>Request body byte count.</summary>
    [JsonIgnore]
    public long BodyLength => _body?.LongLength ?? 0;

    internal byte[]? Buffer => _body;

    internal static string NormalizeMethod(string method) {
        if (string.IsNullOrWhiteSpace(method)) throw new ArgumentException("A request method is required.", nameof(method));
        string normalized = method.ToUpperInvariant();
        if (normalized is not ("GET" or "HEAD" or "POST" or "PUT" or "PATCH" or "DELETE" or "OPTIONS"))
            throw new ArgumentException("Unsupported fetch method.", nameof(method));
        return normalized;
    }

    internal static Dictionary<string, string> NormalizeHeaders(IReadOnlyDictionary<string, string>? headers) {
        if (headers?.Count > 128) throw new ArgumentException("Too many request headers.", nameof(headers));
        var normalized = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        long length = 0;
        foreach (var header in headers ?? new Dictionary<string, string>()) {
            if (string.IsNullOrEmpty(header.Key) || header.Key.Any(c => !char.IsAsciiLetterOrDigit(c) && !"!#$%&'*+-.^_`|~".Contains(c)) ||
                header.Value == null || header.Value.Any(c => c is '\r' or '\n' or '\0' || c > 255) ||
                (length += header.Key.Length + header.Value.Length) > 32768)
                throw new ArgumentException("Invalid or oversized request headers.", nameof(headers));
            if (!ForbiddenHeader(header.Key, header.Value)) normalized.Add(header.Key, header.Value);
        }
        return normalized;
    }

    internal static bool ForbiddenHeader(string name, string value) =>
        name.StartsWith("sec-", StringComparison.OrdinalIgnoreCase) || name.StartsWith("proxy-", StringComparison.OrdinalIgnoreCase) ||
        (name.ToLowerInvariant() is "x-http-method" or "x-http-method-override" or "x-method-override" &&
         value.Split(',').Any(method => method.Trim().ToUpperInvariant() is "CONNECT" or "TRACE" or "TRACK")) ||
        name.ToLowerInvariant() is "accept-charset" or "accept-encoding" or "access-control-request-headers" or
            "access-control-request-method" or "connection" or "content-length" or "cookie" or "cookie2" or "date" or
            "dnt" or "expect" or "host" or "keep-alive" or "origin" or "referer" or "set-cookie" or "te" or "trailer" or
            "transfer-encoding" or "upgrade" or "via";

    private string ComputeIdentity() {
        using var stream = new MemoryStream();
        static void Field(Stream target, string value) {
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            Span<byte> length = stackalloc byte[4];
            BinaryPrimitives.WriteInt32LittleEndian(length, bytes.Length);
            target.Write(length);
            target.Write(bytes);
        }
        Field(stream, HtmlRuntimeResourcePolicy.Key(Url));
        Field(stream, Method);
        Field(stream, Mode);
        Field(stream, Credentials);
        Field(stream, Redirect);
        Span<byte> headerCount = stackalloc byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(headerCount, Headers.Count);
        stream.Write(headerCount);
        foreach (var header in Headers.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)) {
            Field(stream, header.Key.ToLowerInvariant());
            Field(stream, header.Value);
        }
        stream.WriteByte(HasBody ? (byte)1 : (byte)0);
        if (_body != null) stream.Write(_body);
        return Convert.ToHexString(SHA256.HashData(stream.ToArray())).ToLowerInvariant();
    }
}

/// <summary>One occurrence of a dynamic request which the offline worker could not satisfy.</summary>
public sealed class HtmlRuntimeFetchDiscovery {
    /// <summary>Creates a request occurrence.</summary>
    [JsonConstructor]
    public HtmlRuntimeFetchDiscovery(HtmlRuntimeFetchRequest request, int occurrence) {
        Request = request ?? throw new ArgumentNullException(nameof(request));
        if (occurrence <= 0) throw new ArgumentOutOfRangeException(nameof(occurrence));
        Occurrence = occurrence;
    }
    /// <summary>Normalized request.</summary>
    public HtmlRuntimeFetchRequest Request { get; }
    /// <summary>One-based occurrence for repeated identical requests.</summary>
    public int Occurrence { get; }
    /// <summary>Stable identity including the repeated-request occurrence.</summary>
    [JsonIgnore]
    public string Identity => Request.Identity + ":" + Occurrence;
}

/// <summary>One direct HTTP response in an exact dynamic request transcript. A preflight response belongs to the same URL.</summary>
public sealed class HtmlRuntimeFetchHop {
    /// <summary>Creates an immutable response hop and its optional CORS preflight response.</summary>
    [JsonConstructor]
    public HtmlRuntimeFetchHop(HtmlRuntimeResource response, HtmlRuntimeResource? preflightResponse = null) {
        Response = response ?? throw new ArgumentNullException(nameof(response));
        if (Response.RedirectCount != 0 || HtmlRuntimeResourcePolicy.Key(Response.Url) != HtmlRuntimeResourcePolicy.Key(Response.FinalUrl))
            throw new ArgumentException("A fetch hop must contain one direct HTTP response.", nameof(response));
        if (preflightResponse != null && (preflightResponse.RedirectCount != 0 ||
            HtmlRuntimeResourcePolicy.Key(preflightResponse.Url) != HtmlRuntimeResourcePolicy.Key(Response.Url) ||
            HtmlRuntimeResourcePolicy.Key(preflightResponse.FinalUrl) != HtmlRuntimeResourcePolicy.Key(Response.Url)))
            throw new ArgumentException("A preflight response must be direct and use the hop URL.", nameof(preflightResponse));
        PreflightResponse = preflightResponse;
    }

    /// <summary>Direct response for this request URL, including redirect status and Location when applicable.</summary>
    public HtmlRuntimeResource Response { get; }
    /// <summary>Direct OPTIONS response acquired before this hop when CORS preflight was required.</summary>
    public HtmlRuntimeResource? PreflightResponse { get; }
}

/// <summary>An ordered HTTP response transcript bound to one exact dynamic request occurrence.</summary>
public sealed class HtmlRuntimeFetchReplay {
    /// <summary>Creates a one-response replay for requests without redirects or preflight.</summary>
    public HtmlRuntimeFetchReplay(HtmlRuntimeFetchRequest request, int occurrence, HtmlRuntimeResource response)
        : this(request, occurrence, new[] { new HtmlRuntimeFetchHop(response) }) { }

    /// <summary>Creates a bounded, ordered replay from direct HTTP response hops.</summary>
    [JsonConstructor]
    public HtmlRuntimeFetchReplay(HtmlRuntimeFetchRequest request, int occurrence, IReadOnlyList<HtmlRuntimeFetchHop> hops) {
        Request = request ?? throw new ArgumentNullException(nameof(request));
        if (occurrence <= 0) throw new ArgumentOutOfRangeException(nameof(occurrence));
        ArgumentNullException.ThrowIfNull(hops);
        if (hops.Count is < 1 or > 16) throw new ArgumentException("A dynamic replay requires one to sixteen direct response hops.", nameof(hops));
        HtmlRuntimeFetchHop[] retained = hops.ToArray();
        Uri expected = Request.Url;
        for (int index = 0; index < retained.Length; index++) {
            HtmlRuntimeFetchHop hop = retained[index] ?? throw new ArgumentException("A dynamic replay hop cannot be null.", nameof(hops));
            if (HtmlRuntimeResourcePolicy.Key(hop.Response.Url) != HtmlRuntimeResourcePolicy.Key(expected))
                throw new ArgumentException("A dynamic replay hop has a different request URL.", nameof(hops));
            bool redirectStatus = hop.Response.StatusCode is 301 or 302 or 303 or 307 or 308;
            string? location = null;
            bool redirected = redirectStatus && hop.Response.Headers.TryGetValue("Location", out location);
            if (index == retained.Length - 1) {
                if (redirected) throw new ArgumentException("A dynamic replay cannot end at a redirect response.", nameof(hops));
            } else {
                if (!redirected) throw new ArgumentException("A dynamic replay cannot continue after a final response.", nameof(hops));
                expected = new Uri(HtmlRuntimeResourcePolicy.Key(new Uri(expected, location!)));
            }
        }
        Hops = Array.AsReadOnly(retained);
        Occurrence = occurrence;
    }
    /// <summary>Normalized original request.</summary>
    public HtmlRuntimeFetchRequest Request { get; }
    /// <summary>One-based occurrence for repeated identical requests.</summary>
    public int Occurrence { get; }
    /// <summary>Ordered direct responses, including redirect and preflight responses.</summary>
    public IReadOnlyList<HtmlRuntimeFetchHop> Hops { get; }
    /// <summary>Final direct response in the transcript.</summary>
    [JsonIgnore]
    public HtmlRuntimeResource Response => Hops[^1].Response;
    /// <summary>Stable identity including the repeated-request occurrence.</summary>
    [JsonIgnore]
    public string Identity => Request.Identity + ":" + Occurrence;
}
