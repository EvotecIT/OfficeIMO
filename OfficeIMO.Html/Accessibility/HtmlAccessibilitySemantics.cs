using AngleSharp.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Resolves the small, deterministic subset of HTML and ARIA accessibility semantics
/// used by OfficeIMO document conversion pipelines.
/// </summary>
public static class HtmlAccessibilitySemantics {
    private const int MaximumAccessibleNameCharacters = 4096;
    private const int MaximumAriaResolutionDepth = 64;
    private const int MaximumTokenSourceCharacters = 8192;
    private const int MaximumTokens = 256;

    /// <summary>Returns whether an element declares the requested ARIA role.</summary>
    public static bool HasRole(IElement element, string role) =>
        element != null && ContainsToken(element.GetAttribute("role"), role);

    /// <summary>Returns whether an element declares the requested EPUB structural semantic.</summary>
    public static bool HasEpubType(IElement element, string semanticType) {
        if (element == null || string.IsNullOrWhiteSpace(semanticType)) return false;

        string? value = element.GetAttribute("epub:type");
        if (string.IsNullOrWhiteSpace(value)) {
            foreach (IAttr attribute in element.Attributes) {
                if (attribute.Name.Equals("epub:type", StringComparison.OrdinalIgnoreCase)) {
                    value = attribute.Value;
                    break;
                }
            }
        }
        return ContainsToken(value, semanticType);
    }

    /// <summary>
    /// Resolves a heading level from a native heading element or an ARIA heading role.
    /// ARIA levels outside the Markdown heading range are clamped to levels 1 through 6.
    /// </summary>
    public static bool TryGetHeadingLevel(IElement element, out int level) {
        level = 0;
        if (element == null) return false;

        string tagName = element.TagName;
        if (tagName.Length == 2
            && (tagName[0] == 'h' || tagName[0] == 'H')
            && tagName[1] >= '1'
            && tagName[1] <= '6') {
            level = tagName[1] - '0';
            return true;
        }

        if (!HasRole(element, "heading")) return false;
        if (!int.TryParse(
                element.GetAttribute("aria-level"),
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                out level)) {
            level = 2;
        }
        if (level < 1) level = 1;
        if (level > 6) level = 6;
        return true;
    }

    /// <summary>
    /// Resolves an accessible name using ARIA labelling, host-language alternatives,
    /// optional element text, and title fallback. This does not mutate the DOM.
    /// </summary>
    /// <param name="element">Element to name.</param>
    /// <param name="includeTextFallback">Whether normalized descendant text may supply the name.</param>
    public static string GetAccessibleName(IElement element, bool includeTextFallback = false) =>
        GetAccessibleName(element, includeTextFallback, new HtmlAccessibleNameContext());

    internal static string GetAccessibleName(IElement element, bool includeTextFallback,
        HtmlAccessibleNameContext context) =>
        context.Limit(GetAccessibleName(element, includeTextFallback, treatAsImage: false,
            new HashSet<IElement>(), context, depth: 0));

    /// <summary>
    /// Resolves an accessible image name, including image <c>alt</c> semantics for custom
    /// elements that a converter explicitly aliases to an image.
    /// </summary>
    public static string GetImageAccessibleName(IElement element) {
        var context = new HtmlAccessibleNameContext();
        return context.Limit(GetAccessibleName(element,
            includeTextFallback: false, treatAsImage: true, new HashSet<IElement>(),
            context, depth: 0));
    }

    private static string GetAccessibleName(
        IElement element,
        bool includeTextFallback,
        bool treatAsImage,
        ISet<IElement> resolutionPath,
        HtmlAccessibleNameContext context,
        int depth) {
        if (element == null) return string.Empty;
        if (depth > MaximumAriaResolutionDepth) return string.Empty;
        if (!context.TryConsumeTraversalWork()) return string.Empty;
        if (!resolutionPath.Add(element)) return string.Empty;

        try {
            string labelledBy = ResolveLabelledBy(element, resolutionPath, context, depth);
            if (labelledBy.Length > 0) return labelledBy;

            string ariaLabel = context.NormalizeText(element.GetAttribute("aria-label"));
            if (ariaLabel.Length > 0) return ariaLabel;

            string tagName = element.TagName;
            if ((treatAsImage
                 || tagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
                 || tagName.Equals("AREA", StringComparison.OrdinalIgnoreCase))
                && element.HasAttribute("alt")) {
                return context.NormalizeText(element.GetAttribute("alt"));
            }
            if (tagName.Equals("INPUT", StringComparison.OrdinalIgnoreCase)
                && string.Equals(element.GetAttribute("type"), "image", StringComparison.OrdinalIgnoreCase)
                && element.HasAttribute("alt")) {
                return context.NormalizeText(element.GetAttribute("alt"));
            }
            if (tagName.Equals("SVG", StringComparison.OrdinalIgnoreCase)) {
                IElement? titleElement = element.Children.FirstOrDefault(static child =>
                    child.TagName.Equals("TITLE", StringComparison.OrdinalIgnoreCase));
                string svgTitle = context.GetBoundedText(titleElement);
                if (svgTitle.Length > 0) return svgTitle;
            }

            if (includeTextFallback) {
                string text = context.GetBoundedText(element);
                if (text.Length > 0) return text;
            }

            return context.NormalizeText(element.GetAttribute("title"));
        } finally {
            resolutionPath.Remove(element);
        }
    }

    /// <summary>Returns whether an element is explicitly hidden from the accessibility tree.</summary>
    public static bool IsAriaHidden(IElement element) =>
        element != null
        && string.Equals(element.GetAttribute("aria-hidden")?.Trim(), "true", StringComparison.OrdinalIgnoreCase);

    internal static bool ContainsToken(string? value, string token) {
        if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(token)) return false;
        foreach (string candidate in EnumerateTokens(value!)) {
            if (candidate.Equals(token, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static string ResolveLabelledBy(IElement element, ISet<IElement> resolutionPath,
        HtmlAccessibleNameContext context, int depth) {
        string? value = element.GetAttribute("aria-labelledby");
        if (string.IsNullOrWhiteSpace(value) || element.Owner == null) return string.Empty;

        var labels = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        int outputLength = 0;
        foreach (string id in EnumerateTokens(value!)) {
            if (!seen.Add(id)) continue;
            IElement? label = element.Owner.GetElementById(id);
            if (label == null || ReferenceEquals(label, element)) continue;
            string text = GetAccessibleName(label, includeTextFallback: true,
                treatAsImage: false, resolutionPath, context, depth + 1);
            if (text.Length == 0) continue;
            int available = MaximumAccessibleNameCharacters - outputLength - (labels.Count > 0 ? 1 : 0);
            if (available <= 0) break;
            if (text.Length > available) text = text.Substring(0, available).TrimEnd();
            if (text.Length > 0) {
                labels.Add(text);
                outputLength += text.Length + (labels.Count > 1 ? 1 : 0);
            }
        }
        return string.Join(" ", labels);
    }

    private static string NormalizeText(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var builder = new System.Text.StringBuilder(
            Math.Min(value!.Length, MaximumAccessibleNameCharacters));
        bool pendingSpace = false;
        foreach (char character in value) {
            if (IsTokenSeparator(character)) {
                if (builder.Length > 0) pendingSpace = true;
                continue;
            }
            if (pendingSpace && builder.Length < MaximumAccessibleNameCharacters) {
                builder.Append(' ');
            }
            pendingSpace = false;
            if (builder.Length >= MaximumAccessibleNameCharacters) break;
            builder.Append(character);
        }
        return builder.ToString();
    }

    private static IEnumerable<string> EnumerateTokens(string value) {
        int maximum = Math.Min(value.Length, MaximumTokenSourceCharacters);
        int count = 0;
        int offset = 0;
        while (offset < maximum && count < MaximumTokens) {
            while (offset < maximum && IsTokenSeparator(value[offset])) offset++;
            int start = offset;
            while (offset < maximum && !IsTokenSeparator(value[offset])) offset++;
            if (offset > start) {
                count++;
                yield return value.Substring(start, offset - start);
            }
        }
    }

    private static bool IsTokenSeparator(char value) =>
        value == ' ' || value == '\t' || value == '\r' || value == '\n' || value == '\f';

    internal sealed class HtmlAccessibleNameContext {
        private const int MaximumAggregateNameCharacters = 4 * 1024 * 1024;
        private const int MaximumTraversalWork = 4 * 1024 * 1024;
        private int _remainingCharacters = MaximumAggregateNameCharacters;
        private int _remainingTraversalWork = MaximumTraversalWork;

        internal bool TryConsumeTraversalWork(int units = 1) {
            if (units <= 0) units = 1;
            if (_remainingTraversalWork < units) {
                _remainingTraversalWork = 0;
                return false;
            }
            _remainingTraversalWork -= units;
            return true;
        }

        internal string NormalizeText(string? value) {
            if (string.IsNullOrEmpty(value) || !TryConsumeTraversalWork(value!.Length)) {
                return string.Empty;
            }
            return HtmlAccessibilitySemantics.NormalizeText(value);
        }

        internal string Limit(string value) {
            if (_remainingCharacters <= 0 || value.Length == 0) return string.Empty;
            int length = Math.Min(value.Length,
                Math.Min(MaximumAccessibleNameCharacters, _remainingCharacters));
            _remainingCharacters -= length;
            return length == value.Length ? value : value.Substring(0, length).TrimEnd();
        }

        internal string GetBoundedText(IElement? element) {
            if (element == null || _remainingCharacters <= 0 || !TryConsumeTraversalWork()) return string.Empty;
            var builder = new System.Text.StringBuilder(MaximumAccessibleNameCharacters);
            var stack = new Stack<INode>();
            for (int index = element.ChildNodes.Length - 1; index >= 0; index--) {
                stack.Push(element.ChildNodes[index]);
            }
            bool pendingSpace = false;
            while (stack.Count > 0 && builder.Length < MaximumAccessibleNameCharacters) {
                if (!TryConsumeTraversalWork()) break;
                INode node = stack.Pop();
                if (node.NodeType == NodeType.Text) {
                    string? sourceText = node.TextContent;
                    if (string.IsNullOrEmpty(sourceText) || !TryConsumeTraversalWork(sourceText!.Length)) {
                        continue;
                    }
                    string text = HtmlAccessibilitySemantics.NormalizeText(sourceText);
                    if (text.Length == 0) continue;
                    if (pendingSpace && builder.Length < MaximumAccessibleNameCharacters) builder.Append(' ');
                    int available = MaximumAccessibleNameCharacters - builder.Length;
                    builder.Append(text, 0, Math.Min(text.Length, available));
                    pendingSpace = true;
                    continue;
                }
                for (int index = node.ChildNodes.Length - 1; index >= 0; index--) {
                    stack.Push(node.ChildNodes[index]);
                }
            }
            return builder.ToString().TrimEnd();
        }
    }
}
