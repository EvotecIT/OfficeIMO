namespace OfficeIMO.Markdown;

/// <summary>
/// Context supplied to visual-host HTML round-trip hints.
/// </summary>
public sealed class MarkdownVisualElementRoundTripContext {
    internal MarkdownVisualElementRoundTripContext(
        string elementName,
        MarkdownVisualElement visualElement,
        string payload,
        string? caption) {
        ElementName = elementName ?? string.Empty;
        VisualElement = visualElement ?? throw new ArgumentNullException(nameof(visualElement));
        Payload = payload ?? string.Empty;
        Caption = caption;
        FenceInfoString = BuildFenceInfoString(visualElement);
    }

    /// <summary>Original HTML element name that carried the shared visual contract.</summary>
    public string ElementName { get; }

    /// <summary>Parsed shared visual metadata descriptor.</summary>
    public MarkdownVisualElement VisualElement { get; }

    /// <summary>Decoded raw payload content.</summary>
    public string Payload { get; }

    /// <summary>Optional caption recovered from the HTML host, for example a sibling <c>figcaption</c>.</summary>
    public string? Caption { get; }

    /// <summary>Default reconstructed fence info string for the parsed visual host.</summary>
    public string FenceInfoString { get; }

    /// <summary>
    /// Creates a semantic fenced block from the parsed visual host, optionally overriding semantic kind,
    /// fence info string, payload, or caption.
    /// </summary>
    public SemanticFencedBlock CreateBlock(
        string? semanticKind = null,
        string? fenceInfoString = null,
        string? payload = null,
        string? caption = null) {
        return new SemanticFencedBlock(
            string.IsNullOrWhiteSpace(semanticKind) ? VisualElement.VisualKind : semanticKind!,
            string.IsNullOrWhiteSpace(fenceInfoString) ? FenceInfoString : fenceInfoString!,
            payload ?? Payload,
            caption ?? Caption);
    }

    internal static string BuildFenceInfoString(MarkdownVisualElement visualElement) {
        if (visualElement == null || string.IsNullOrWhiteSpace(visualElement.FenceLanguage)) {
            return string.Empty;
        }

        var fenceAdditionalInfo = visualElement.FenceAdditionalInfo;
        if (!string.IsNullOrWhiteSpace(fenceAdditionalInfo)) {
            return visualElement.FenceLanguage + " " + fenceAdditionalInfo!.Trim();
        }

        var parts = new List<string> {
            visualElement.FenceLanguage
        };

        if (!string.IsNullOrWhiteSpace(visualElement.FenceElementId)) {
            parts.Add("#" + visualElement.FenceElementId);
        }

        var fenceClasses = visualElement.FenceClasses;
        for (int i = 0; i < fenceClasses.Count; i++) {
            parts.Add("." + fenceClasses[i]);
        }

        if (!string.IsNullOrWhiteSpace(visualElement.VisualTitle)) {
            parts.Add("title=" + QuoteFenceAttributeValue(visualElement.VisualTitle!));
        }

        return string.Join(" ", parts);
    }

    private static string QuoteFenceAttributeValue(string value) {
        var normalized = value ?? string.Empty;
        var escaped = normalized
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"");
        return "\"" + escaped + "\"";
    }
}

/// <summary>
/// Ordered host/plugin seam for reinterpreting shared visual HTML contract elements during HTML-to-markdown conversion.
/// </summary>
public sealed class MarkdownVisualElementRoundTripHint {
    private readonly Func<MarkdownVisualElementRoundTripContext, SemanticFencedBlock?> _transform;

    /// <summary>Create a new visual-host HTML round-trip hint.</summary>
    public MarkdownVisualElementRoundTripHint(
        string id,
        string name,
        Func<MarkdownVisualElementRoundTripContext, SemanticFencedBlock?> transform) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new ArgumentException("Hint id is required.", nameof(id));
        }

        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Hint name is required.", nameof(name));
        }

        Id = id.Trim();
        Name = name.Trim();
        _transform = transform ?? throw new ArgumentNullException(nameof(transform));
    }

    /// <summary>Stable hint identifier used for deduplication and diagnostics.</summary>
    public string Id { get; }

    /// <summary>Friendly hint name used for diagnostics and documentation.</summary>
    public string Name { get; }

    /// <summary>Attempts to create a semantic fenced block from the supplied visual-host context.</summary>
    public SemanticFencedBlock? TryCreateBlock(MarkdownVisualElementRoundTripContext context) {
        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        return _transform(context);
    }
}
