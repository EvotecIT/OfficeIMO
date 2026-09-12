using System.Text;
using System.Text.Json.Serialization;

namespace OfficeIMO.Html.Runtime;

/// <summary>An immutable resource supplied to a session or retained with a capture.</summary>
public sealed class HtmlRuntimeResource {
    private readonly byte[] _content;

    /// <summary>Creates a resource with an absolute HTTP(S) identity and an independent byte snapshot.</summary>
    [JsonConstructor]
    public HtmlRuntimeResource(Uri url, byte[] content, string contentType, int statusCode = 200, Uri? finalUrl = null, int redirectCount = 0) {
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
    /// <summary>Encoded response size.</summary>
    [JsonIgnore]
    public long Length => _content.LongLength;
    internal byte[] Buffer => _content;
}
