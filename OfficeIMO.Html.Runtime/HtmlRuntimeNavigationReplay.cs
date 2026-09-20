using System.Buffers.Binary;
using System.Collections.ObjectModel;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Serialization;

namespace OfficeIMO.Html.Runtime;

/// <summary>Reason a top-level document request replaces the active document.</summary>
public enum HtmlRuntimeNavigationKind {
    /// <summary>A link, form, script, automation, or host requested a new document.</summary>
    Navigate,
    /// <summary>The active history entry is being loaded again.</summary>
    Reload,
    /// <summary>Back or forward traversal selected an entry owned by another document.</summary>
    Traverse
}

/// <summary>An immutable top-level document request that can be acquired once and replayed without worker network access.</summary>
public sealed class HtmlRuntimeNavigationRequest {
    private readonly byte[]? _body;

    /// <summary>Creates an exact document request occurrence envelope and infers body presence from <paramref name="body"/>.</summary>
    public HtmlRuntimeNavigationRequest(Uri url, Uri initiatorUrl, Uri? referrer = null, string method = "GET",
        IReadOnlyDictionary<string, string>? headers = null, byte[]? body = null,
        HtmlRuntimeNavigationKind kind = HtmlRuntimeNavigationKind.Navigate, bool replaceHistoryEntry = false,
        int historyEntryIndex = -1)
        : this(url, initiatorUrl, referrer, method, headers, body, body != null, kind, replaceHistoryEntry,
            historyEntryIndex) { }

    /// <summary>Creates an exact document request occurrence envelope.</summary>
    [JsonConstructor]
    public HtmlRuntimeNavigationRequest(Uri url, Uri initiatorUrl, Uri? referrer, string method,
        IReadOnlyDictionary<string, string>? headers, byte[]? body, bool hasBody,
        HtmlRuntimeNavigationKind kind, bool replaceHistoryEntry, int historyEntryIndex) {
        Url = HtmlRuntimeResourcePolicy.ValidateUrl(url);
        InitiatorUrl = HtmlRuntimeResourcePolicy.ValidateUrl(initiatorUrl);
        Referrer = referrer == null ? null : WithoutFragment(HtmlRuntimeResourcePolicy.ValidateUrl(referrer));
        Method = HtmlRuntimeFetchRequest.NormalizeMethod(method);
        if (Method is not ("GET" or "POST"))
            throw new ArgumentException("Document navigation supports GET and POST requests.", nameof(method));
        HasBody = hasBody;
        if (HasBody && Method == "GET")
            throw new ArgumentException("GET navigation requests cannot have a body.", nameof(body));
        Headers = new ReadOnlyDictionary<string, string>(HtmlRuntimeFetchRequest.NormalizeHeaders(headers));
        _body = HasBody ? (byte[])(body ?? Array.Empty<byte>()).Clone() : null;
        if (!Enum.IsDefined(kind)) throw new ArgumentOutOfRangeException(nameof(kind));
        if (historyEntryIndex < -1) throw new ArgumentOutOfRangeException(nameof(historyEntryIndex));
        if (kind == HtmlRuntimeNavigationKind.Navigate && historyEntryIndex != -1)
            throw new ArgumentException("A new navigation cannot select an existing history entry.", nameof(historyEntryIndex));
        if (kind != HtmlRuntimeNavigationKind.Navigate && historyEntryIndex < 0)
            throw new ArgumentException("Reload and traversal requests require a history entry index.", nameof(historyEntryIndex));
        Kind = kind;
        ReplaceHistoryEntry = replaceHistoryEntry;
        HistoryEntryIndex = historyEntryIndex;
        Identity = ComputeIdentity();
    }

    /// <summary>Originally requested absolute document URL.</summary>
    public Uri Url { get; }
    /// <summary>Document URL whose browsing context initiated the request.</summary>
    public Uri InitiatorUrl { get; }
    /// <summary>Referrer sent for the first request after fragment removal and default cross-origin reduction.</summary>
    public Uri? Referrer { get; }
    /// <summary>Uppercase HTTP method.</summary>
    public string Method { get; }
    /// <summary>Normalized outgoing headers after browser-forbidden names were removed.</summary>
    public IReadOnlyDictionary<string, string> Headers { get; }
    /// <summary>Independent request body bytes, or <see langword="null"/>.</summary>
    public byte[]? Body => _body == null ? null : (byte[])_body.Clone();
    /// <summary>Whether the request has a body, including a zero-byte body.</summary>
    public bool HasBody { get; }
    /// <summary>Browser-history operation that caused the document load.</summary>
    public HtmlRuntimeNavigationKind Kind { get; }
    /// <summary>Whether a new navigation replaces the current history entry.</summary>
    public bool ReplaceHistoryEntry { get; }
    /// <summary>Selected history entry for reload or traversal, otherwise -1.</summary>
    public int HistoryEntryIndex { get; }
    /// <summary>Stable lowercase SHA-256 identity over the complete navigation envelope.</summary>
    public string Identity { get; }
    /// <summary>Request body byte count.</summary>
    [JsonIgnore]
    public long BodyLength => _body?.LongLength ?? 0;

    internal byte[]? Buffer => _body;

    internal static Uri? DefaultReferrer(Uri initiator, Uri target) {
        initiator = HtmlRuntimeResourcePolicy.ValidateUrl(initiator);
        target = HtmlRuntimeResourcePolicy.ValidateUrl(target);
        if (initiator.Scheme == Uri.UriSchemeHttps && target.Scheme != Uri.UriSchemeHttps) return null;
        return HtmlRuntimeResourcePolicy.Origin(initiator) == HtmlRuntimeResourcePolicy.Origin(target)
            ? WithoutFragment(initiator)
            : new Uri(HtmlRuntimeResourcePolicy.Origin(initiator) + "/");
    }

    internal static Uri WithoutFragment(Uri value) => new UriBuilder(value) { Fragment = string.Empty }.Uri;

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
        Field(stream, HtmlRuntimeResourcePolicy.Key(InitiatorUrl));
        Field(stream, Referrer?.AbsoluteUri ?? string.Empty);
        Field(stream, Method);
        Field(stream, Kind.ToString());
        Field(stream, ReplaceHistoryEntry ? "replace" : "append");
        Field(stream, HistoryEntryIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
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

/// <summary>One occurrence of a top-level document request which the offline worker could not satisfy.</summary>
public sealed class HtmlRuntimeNavigationDiscovery {
    /// <summary>Creates a document-request occurrence.</summary>
    [JsonConstructor]
    public HtmlRuntimeNavigationDiscovery(HtmlRuntimeNavigationRequest request, int occurrence) {
        Request = request ?? throw new ArgumentNullException(nameof(request));
        if (occurrence <= 0) throw new ArgumentOutOfRangeException(nameof(occurrence));
        Occurrence = occurrence;
    }
    /// <summary>Normalized navigation request.</summary>
    public HtmlRuntimeNavigationRequest Request { get; }
    /// <summary>One-based occurrence for repeated identical requests.</summary>
    public int Occurrence { get; }
    /// <summary>Stable identity including the repeated-request occurrence.</summary>
    [JsonIgnore]
    public string Identity => Request.Identity + ":" + Occurrence;
}

/// <summary>One direct response in a top-level document-navigation transcript.</summary>
public sealed class HtmlRuntimeNavigationHop {
    /// <summary>Creates one immutable direct response hop.</summary>
    [JsonConstructor]
    public HtmlRuntimeNavigationHop(HtmlRuntimeResource response) {
        Response = response ?? throw new ArgumentNullException(nameof(response));
        if (Response.RedirectCount != 0 || HtmlRuntimeResourcePolicy.Key(Response.Url) != HtmlRuntimeResourcePolicy.Key(Response.FinalUrl))
            throw new ArgumentException("A navigation hop must contain one direct HTTP response.", nameof(response));
    }
    /// <summary>Direct response for this request URL, including redirect status and Location when applicable.</summary>
    public HtmlRuntimeResource Response { get; }
}

/// <summary>An ordered HTTP response transcript bound to one exact top-level navigation occurrence.</summary>
public sealed class HtmlRuntimeNavigationReplay {
    /// <summary>Creates a one-response replay without redirects.</summary>
    public HtmlRuntimeNavigationReplay(HtmlRuntimeNavigationRequest request, int occurrence, HtmlRuntimeResource response)
        : this(request, occurrence, new[] { new HtmlRuntimeNavigationHop(response) }) { }

    /// <summary>Creates a bounded ordered replay from direct HTTP response hops.</summary>
    [JsonConstructor]
    public HtmlRuntimeNavigationReplay(HtmlRuntimeNavigationRequest request, int occurrence,
        IReadOnlyList<HtmlRuntimeNavigationHop> hops) {
        Request = request ?? throw new ArgumentNullException(nameof(request));
        if (occurrence <= 0) throw new ArgumentOutOfRangeException(nameof(occurrence));
        ArgumentNullException.ThrowIfNull(hops);
        if (hops.Count is < 1 or > 16)
            throw new ArgumentException("A navigation replay requires one to sixteen direct response hops.", nameof(hops));
        HtmlRuntimeNavigationHop[] retained = hops.ToArray();
        Uri expected = Request.Url;
        for (int index = 0; index < retained.Length; index++) {
            HtmlRuntimeNavigationHop hop = retained[index]
                ?? throw new ArgumentException("A navigation replay hop cannot be null.", nameof(hops));
            if (HtmlRuntimeResourcePolicy.Key(hop.Response.Url) != HtmlRuntimeResourcePolicy.Key(expected))
                throw new ArgumentException("A navigation replay hop has a different request URL.", nameof(hops));
            bool redirectStatus = hop.Response.StatusCode is 301 or 302 or 303 or 307 or 308;
            string? location = null;
            bool redirected = redirectStatus && hop.Response.Headers.TryGetValue("Location", out location);
            if (index == retained.Length - 1) {
                if (redirected) throw new ArgumentException("A navigation replay cannot end at a redirect response.", nameof(hops));
            } else {
                if (!redirected) throw new ArgumentException("A navigation replay cannot continue after a final response.", nameof(hops));
                expected = new Uri(HtmlRuntimeResourcePolicy.Key(new Uri(expected, location!)));
            }
        }
        Hops = Array.AsReadOnly(retained);
        Occurrence = occurrence;
    }

    /// <summary>Normalized original document request.</summary>
    public HtmlRuntimeNavigationRequest Request { get; }
    /// <summary>One-based occurrence for repeated identical document requests.</summary>
    public int Occurrence { get; }
    /// <summary>Ordered direct redirect and final responses.</summary>
    public IReadOnlyList<HtmlRuntimeNavigationHop> Hops { get; }
    /// <summary>Final direct response.</summary>
    [JsonIgnore]
    public HtmlRuntimeResource Response => Hops[^1].Response;
    /// <summary>Stable identity including the repeated-request occurrence.</summary>
    [JsonIgnore]
    public string Identity => Request.Identity + ":" + Occurrence;
}
