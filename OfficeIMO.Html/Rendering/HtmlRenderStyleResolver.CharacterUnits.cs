namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    // This is only a fast path for deciding whether to obtain font metrics;
    // the length parser, not this substring test, validates dimension tokens.
    private static bool ContainsCharacterUnit(string value) => value.IndexOf("ch", StringComparison.OrdinalIgnoreCase) >= 0;
}
