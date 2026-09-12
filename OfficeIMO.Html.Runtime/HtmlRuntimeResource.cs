using System.Text;
using System.Text.Json.Serialization;

namespace OfficeIMO.Html.Runtime;

/// <summary>An immutable resource supplied to a session or retained with a capture.</summary>
public sealed class HtmlRuntimeResource {
    private readonly byte[] _content;

    /// <summary>Creates a resource with an absolute HTTP(S) identity and an independent byte snapshot.</summary>
    [JsonConstructor]
    public HtmlRuntimeResource(Uri url, byte[] content, string contentType, int statusCode = 200, Uri? finalUrl = null, int redirectCount = 0,
        IReadOnlyDictionary<string, string>? headers = null, string statusText = "") {
        Url = HtmlRuntimeResourcePolicy.ValidateUrl(url);
        FinalUrl = HtmlRuntimeResourcePolicy.ValidateUrl(finalUrl ?? url);
        if (redirectCount < 0) throw new ArgumentOutOfRangeException(nameof(redirectCount));
        RedirectCount = redirectCount;
        ArgumentNullException.ThrowIfNull(content);
        if (string.IsNullOrWhiteSpace(contentType) || contentType.Length > 256 || contentType.IndexOfAny(new[] { '\r', '\n' }) >= 0)
            throw new ArgumentException("A valid resource content type is required.", nameof(contentType));
        if (statusCode < 200 || statusCode > 599) throw new ArgumentOutOfRangeException(nameof(statusCode));
        _content = (byte[])content.Clone();
        ContentType = contentType;
        StatusCode = statusCode;
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        long headerBytes = 0;
        foreach (var header in headers ?? new Dictionary<string, string>()) {
            if (values.Count >= 128 || string.IsNullOrEmpty(header.Key) || header.Key.Any(c => !char.IsAsciiLetterOrDigit(c) && !"!#$%&'*+-.^_`|~".Contains(c)) ||
                header.Value == null || header.Value.IndexOfAny(new[] { '\r', '\n', '\0' }) >= 0 || (headerBytes += header.Key.Length + header.Value.Length) > 32768)
                throw new ArgumentException("Invalid or oversized resource headers.", nameof(headers));
            values.Add(header.Key, header.Value);
        }
        Headers = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(values);
        if (statusText == null || statusText.Length > 1024 || statusText.Any(c => c < 32 || c == 127)) throw new ArgumentException("Invalid status text.", nameof(statusText));
        StatusText = statusText;
    }

    /// <summary>Creates a UTF-8 resource, such as a script, stylesheet or JSON document.</summary>
    public static HtmlRuntimeResource FromText(Uri url, string content, string contentType) => new(url, Encoding.UTF8.GetBytes(content), contentType);
    /// <summary>Absolute resource identity. URL paths and queries are case-sensitive.</summary>
    public Uri Url { get; }
    /// <summary>Final response identity after redirects.</summary>
    public Uri FinalUrl { get; }
    /// <summary>Number of redirects followed.</summary>
    public int RedirectCount { get; }
    /// <summary>Independent copy of the response bytes.</summary>
    public byte[] Content => (byte[])_content.Clone();
    /// <summary>Declared response media type, optionally including a charset.</summary>
    public string ContentType { get; }
    /// <summary>HTTP response status.</summary>
    public int StatusCode { get; }
    /// <summary>Immutable response headers. Browser fetch applies its own exposure rules.</summary>
    public IReadOnlyDictionary<string, string> Headers { get; }
    /// <summary>HTTP status text, when supplied by the transport.</summary>
    public string StatusText { get; }
    /// <summary>Encoded response size.</summary>
    [JsonIgnore]
    public long Length => _content.LongLength;
    internal byte[] Buffer => _content;
}
