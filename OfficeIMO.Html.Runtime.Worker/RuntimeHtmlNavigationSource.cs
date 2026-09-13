using System.Net.Http.Headers;
using System.Text;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeHtmlNavigationSource {
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    internal static string Decode(HtmlRuntimeResource resource, int maximumCharacters) {
        if (resource.Length > (long)maximumCharacters * 4 + 3)
            throw new HtmlScriptRuntimeException("The navigated document exceeds MaxInputCharacters.");
        if (!MediaTypeHeaderValue.TryParse(resource.ContentType, out var mediaType))
            throw new HtmlScriptRuntimeException("The navigation response has an invalid Content-Type.");
        string? charset = mediaType.CharSet?.Trim('"');
        if (!string.IsNullOrEmpty(charset) && !charset.Equals("utf-8", StringComparison.OrdinalIgnoreCase) && !charset.Equals("utf8", StringComparison.OrdinalIgnoreCase))
            throw new HtmlScriptRuntimeException("WebApplicationV1 navigation currently requires UTF-8 HTML.");
        string source;
        try { source = StrictUtf8.GetString(resource.Buffer); }
        catch (DecoderFallbackException) { throw new HtmlScriptRuntimeException("The navigation response is not valid UTF-8 HTML."); }
        if (source.Length != 0 && source[0] == '\uFEFF') source = source[1..];
        if (source.Length > maximumCharacters) throw new HtmlScriptRuntimeException("The navigated document exceeds MaxInputCharacters.");
        return source;
    }
}
