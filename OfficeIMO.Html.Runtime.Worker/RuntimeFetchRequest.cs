namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeFetchRequest {
    public string Url { get; set; } = string.Empty;
    public string Method { get; set; } = "GET";
    public Dictionary<string, string> Headers { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public string? Body { get; set; }
    public string Mode { get; set; } = "cors";
    public string Credentials { get; set; } = "same-origin";
    public string Redirect { get; set; } = "follow";

    internal byte[]? Validate(HtmlRuntimeResourcePolicy policy) {
        if (Method is not ("GET" or "HEAD" or "POST" or "PUT" or "PATCH" or "DELETE" or "OPTIONS")) throw new HtmlScriptRuntimeException("Unsupported fetch method.");
        if (Mode is not ("cors" or "same-origin") || Credentials is not ("omit" or "same-origin") || Redirect is not ("follow" or "error"))
            throw new HtmlScriptRuntimeException("Unsupported fetch mode, credentials or redirect option.");
        if (Body != null && Method is "GET" or "HEAD") throw new HtmlScriptRuntimeException("GET and HEAD requests cannot have a body.");
        if (Body?.Length > ((policy.MaxRequestBytes + 2) / 3) * 4) throw new HtmlScriptRuntimeException("Fetch request body exceeds its byte budget.");
        byte[]? body = Body == null ? null : Convert.FromBase64String(Body);
        if (body?.LongLength > policy.MaxRequestBytes) throw new HtmlScriptRuntimeException("Fetch request body exceeds its byte budget.");
        if (Headers == null || Headers.Count > 128) throw new HtmlScriptRuntimeException("Invalid fetch headers.");
        long length = 0;
        var normalized = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var header in Headers) {
            if (string.IsNullOrEmpty(header.Key) || header.Key.Any(c => !char.IsAsciiLetterOrDigit(c) && !"!#$%&'*+-.^_`|~".Contains(c)) ||
                header.Value == null || header.Value.Any(c => c == '\r' || c == '\n' || c == '\0' || c > 255) || (length += header.Key.Length + header.Value.Length) > 32768)
                throw new HtmlScriptRuntimeException("Invalid or oversized fetch headers.");
            if (!ForbiddenHeader(header.Key, header.Value)) normalized.Add(header.Key, header.Value);
        }
        Headers = normalized;
        return body;
    }

    private static bool ForbiddenHeader(string name, string value) => name.StartsWith("sec-", StringComparison.OrdinalIgnoreCase) || name.StartsWith("proxy-", StringComparison.OrdinalIgnoreCase) ||
        (name.ToLowerInvariant() is "x-http-method" or "x-http-method-override" or "x-method-override" && value.Split(',').Any(method => method.Trim().ToUpperInvariant() is "CONNECT" or "TRACE" or "TRACK")) ||
        name.ToLowerInvariant() is "accept-charset" or "accept-encoding" or "access-control-request-headers" or "access-control-request-method" or "connection" or "content-length" or "cookie" or "cookie2" or "date" or "dnt" or "expect" or "host" or "keep-alive" or "origin" or "referer" or "set-cookie" or "te" or "trailer" or "transfer-encoding" or "upgrade" or "via";
}
