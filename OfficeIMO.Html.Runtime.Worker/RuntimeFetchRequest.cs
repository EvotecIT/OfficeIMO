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
        Method = HtmlRuntimeFetchRequest.NormalizeMethod(Method);
        if (Mode is not ("cors" or "same-origin") || Credentials is not ("omit" or "same-origin") || Redirect is not ("follow" or "error"))
            throw new HtmlScriptRuntimeException("Unsupported fetch mode, credentials or redirect option.");
        if (Body != null && Method is "GET" or "HEAD") throw new HtmlScriptRuntimeException("GET and HEAD requests cannot have a body.");
        if (Body?.Length > ((policy.MaxRequestBytes + 2) / 3) * 4) throw new HtmlScriptRuntimeException("Fetch request body exceeds its byte budget.");
        byte[]? body = Body == null ? null : Convert.FromBase64String(Body);
        if (body?.LongLength > policy.MaxRequestBytes) throw new HtmlScriptRuntimeException("Fetch request body exceeds its byte budget.");
        try { Headers = HtmlRuntimeFetchRequest.NormalizeHeaders(Headers); }
        catch (ArgumentException error) { throw new HtmlScriptRuntimeException(error.Message); }
        return body;
    }

    internal HtmlRuntimeFetchRequest ReplayRequest(Uri url, byte[]? body) =>
        new(url, Method, Headers, body, Mode, Credentials, Redirect, Body != null);
}
