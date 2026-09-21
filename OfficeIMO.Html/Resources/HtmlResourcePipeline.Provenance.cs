using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private const char CssCommentMask = '\u0001';
    internal static bool IsProvenanceStyleElement(IElement element) => IsCssStyleElement(element);
    internal static HtmlProvenanceCssScope CollectProvenanceCssImageScope(
        IHtmlDocument document,
        long maximumStylesheetBytes,
        long maximumExpandedBytes) {
        List<HtmlProvenanceDataStylesheet> dataStylesheets = MaterializeDataStylesheets(
            document, maximumStylesheetBytes, maximumExpandedBytes, out long decodedStylesheetBytes);
        try {
            var result = new HtmlProvenanceCssScope { DecodedStylesheetBytes = decodedStylesheetBytes };
            foreach (HtmlProvenanceDataStylesheet stylesheet in dataStylesheets) {
                result.DataStylesheets.Add(stylesheet.Link, stylesheet);
            }
            return result;
        } finally {
            foreach (HtmlProvenanceDataStylesheet stylesheet in dataStylesheets) stylesheet.MaterializedStyle.Remove();
        }
    }

    private static List<HtmlProvenanceDataStylesheet> MaterializeDataStylesheets(
        IHtmlDocument document,
        long maximumStylesheetBytes,
        long maximumExpandedBytes,
        out long decodedStylesheetBytes) {
        var stylesheets = new List<HtmlProvenanceDataStylesheet>();
        long expandedBytes = 0;
        decodedStylesheetBytes = 0;
        try {
            foreach (IElement link in document.QuerySelectorAll("link[href]")) {
                string href = link.GetAttribute("href") ?? string.Empty;
                int commaIndex = href.IndexOf(',');
                int fragmentIndex = commaIndex >= 0 ? href.IndexOf('#', commaIndex + 1) : -1;
                string fragment = fragmentIndex >= 0 ? href.Substring(fragmentIndex) : string.Empty;
                string dataSource = fragmentIndex >= 0 ? href.Substring(0, fragmentIndex) : href;
                if (!IsHtmlStylesheetLink(link) ||
                    !HtmlDataUri.TryParse(dataSource, out HtmlDataUri dataUri) ||
                    !string.Equals(dataUri.MediaType, "text/css", StringComparison.OrdinalIgnoreCase)) continue;

                long decodedByteCount;
                string css;
                try {
                    decodedByteCount = dataUri.EstimateDecodedByteCount();
                    if (decodedByteCount > maximumStylesheetBytes) {
                        throw new InvalidDataException("An embedded HTML stylesheet exceeds the configured asset limit.");
                    }
                    expandedBytes = checked(expandedBytes + decodedByteCount);
                    if (expandedBytes > maximumExpandedBytes) {
                        throw new InvalidDataException("Embedded HTML stylesheets exceed the configured expanded-container limit.");
                    }
                    css = dataUri.DecodeText();
                    decodedStylesheetBytes = expandedBytes;
                } catch (OverflowException exception) {
                    throw new InvalidDataException("Embedded HTML stylesheets declare an invalid expanded size.", exception);
                } catch (UriFormatException) {
                    continue;
                } catch (FormatException) {
                    continue;
                } catch (ArgumentException) {
                    continue;
                }
                if (string.IsNullOrWhiteSpace(css) || link.Parent == null) continue;

                IElement style = document.CreateElement("style");
                style.TextContent = css;
                string media = link.GetAttribute("media") ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(media)) style.SetAttribute("media", media);
                INode parent = link.Parent;
                INode? next = link.NextSibling;
                if (next == null) parent.AppendChild(style);
                else parent.InsertBefore(style, next);
                stylesheets.Add(new HtmlProvenanceDataStylesheet(
                    link, style, css, dataUri.Metadata, fragment));
            }
            return stylesheets;
        } catch {
            foreach (HtmlProvenanceDataStylesheet stylesheet in stylesheets) stylesheet.MaterializedStyle.Remove();
            throw;
        }
    }

    private static bool IsHtmlStylesheetLink(IElement link) {
        if (!IsHtmlNamespaceElement(link)) return false;
        bool stylesheet = (link.GetAttribute("rel") ?? string.Empty)
            .Split(new[] { '\t', '\n', '\f', '\r', ' ' }, StringSplitOptions.RemoveEmptyEntries)
            .Any(token => token.Equals("stylesheet", StringComparison.OrdinalIgnoreCase));
        if (!stylesheet) return false;
        string type = link.GetAttribute("type") ?? string.Empty;
        int parameter = type.IndexOf(';');
        if (parameter >= 0) type = type.Substring(0, parameter);
        type = type.Trim(' ', '\t', '\n', '\f', '\r');
        return type.Length == 0 || type.Equals("text/css", StringComparison.OrdinalIgnoreCase);
    }

    internal static IEnumerable<HtmlCssImageReference> EnumerateProvenanceCssImageReferences(string css) {
        if (string.IsNullOrWhiteSpace(css)) yield break;
        string masked = MaskCssComments(css);
        var emittedRanges = new HashSet<(int Start, int Length)>();
        foreach (Match match in CssUrlExpression.Matches(masked)) {
            if (!IsValidCssUrlMatch(masked, match)) continue;
            if (!IsCssFunctionNameAt(masked, match.Index, "url") ||
                IsInsideCssString(masked, match.Index) ||
                IsImportAtRuleUrl(masked, match.Index) ||
                IsAtRulePreludeUrl(masked, match.Index)) continue;
            Group sourceGroup = match.Groups["url"];
            int leading = 0;
            while (leading < sourceGroup.Length && IsCssWhitespace(sourceGroup.Value[leading])) leading++;
            int trailing = sourceGroup.Length;
            while (trailing > leading && IsCssWhitespace(sourceGroup.Value[trailing - 1])) trailing--;
            if (trailing == leading) continue;
            string source = DecodeCssEscapes(sourceGroup.Value.Substring(leading, trailing - leading));
            var range = (sourceGroup.Index + leading, trailing - leading);
            if (emittedRanges.Add(range)) yield return new HtmlCssImageReference(range.Item1, range.Item2, source);
        }

        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(masked)) {
            if (!emittedRanges.Add((reference.SourceStart, reference.Source.Length))) continue;
            yield return new HtmlCssImageReference(reference.SourceStart, reference.Source.Length, DecodeCssEscapes(reference.Source));
        }
    }

    private static string MaskCssComments(string css) {
        var result = new StringBuilder(css);
        char quote = '\0';
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current is '"' or '\'') { quote = current; continue; }
            if (current != '/' || index + 1 >= css.Length || css[index + 1] != '*') continue;
            result[index++] = CssCommentMask;
            result[index] = CssCommentMask;
            while (index + 1 < css.Length && !(css[index] == '*' && css[index + 1] == '/')) result[++index] = CssCommentMask;
            if (index + 1 < css.Length) { result[index] = CssCommentMask; result[++index] = CssCommentMask; }
        }
        return result.ToString();
    }
}

internal sealed class HtmlProvenanceCssScope {
    internal Dictionary<IElement, HtmlProvenanceDataStylesheet> DataStylesheets { get; } = new();
    internal long DecodedStylesheetBytes { get; set; }

}

internal sealed class HtmlProvenanceDataStylesheet {
    internal HtmlProvenanceDataStylesheet(
        IElement link,
        IElement materializedStyle,
        string css,
        string metadata,
        string fragment) {
        Link = link;
        MaterializedStyle = materializedStyle;
        Css = css;
        Metadata = metadata;
        Fragment = fragment;
    }

    internal IElement Link { get; }
    internal IElement MaterializedStyle { get; }
    internal string Css { get; }
    internal string Metadata { get; }
    internal string Fragment { get; }
}

internal readonly struct HtmlCssImageReference {
    internal HtmlCssImageReference(int start, int length, string value) {
        Start = start;
        Length = length;
        Value = value;
    }

    internal int Start { get; }
    internal int Length { get; }
    internal string Value { get; }
}
