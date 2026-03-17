using OfficeIMO.MarkdownRenderer;

namespace OfficeIMO.MarkdownRenderer.IntelligenceX;

/// <summary>
/// IntelligenceX fence option schemas layered on top of the generic renderer contract.
/// </summary>
public static class IntelligenceXVisualFenceSchemas {
    /// <summary>
    /// Shared IntelligenceX visual option schema for chart, network, and dataview fences.
    /// </summary>
    public static MarkdownFenceOptionSchema Visuals { get; } = new MarkdownFenceOptionSchema(
        "officeimo.intelligencex.visual-options",
        "IntelligenceX Visual Options",
        new[] { "ix-chart", "ix-network", "ix-dataview" },
        new[] {
            MarkdownFenceOptionDefinition.Boolean(
                "pinned",
                aliases: new[] { "pin" }),
            MarkdownFenceOptionDefinition.String(
                "theme",
                aliases: new[] { "palette", "colorScheme" }),
            MarkdownFenceOptionDefinition.String(
                "variant",
                aliases: new[] { "style" }),
            MarkdownFenceOptionDefinition.String(
                "view",
                aliases: new[] { "mode" }),
            MarkdownFenceOptionDefinition.Int32(
                "maxItems",
                aliases: new[] { "max-items", "limit" },
                validator: rawValue => int.TryParse(rawValue, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var parsed)
                    && parsed > 0
                    ? null
                    : "Expected a positive integer value.")
        });
}
