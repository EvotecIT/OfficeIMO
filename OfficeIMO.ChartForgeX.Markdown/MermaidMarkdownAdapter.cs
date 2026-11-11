using System.Text;
using global::ChartForgeX.Diagnostics;
using global::ChartForgeX.Markup;
using global::ChartForgeX.Markup.Mermaid;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;

namespace OfficeIMO.ChartForgeX.Markdown;

/// <summary>Image payload used when Mermaid fences are replaced by rendered diagrams.</summary>
public enum MermaidMarkdownImageFormat {
    /// <summary>Embedded SVG, the default. PDF draws it as vector content, and Word stores the SVG with a PNG fallback picture.</summary>
    Svg,
    /// <summary>Embedded PNG, for consumers that handle SVG poorly, such as HTML email or document import services.</summary>
    Png
}

/// <summary>Connects ChartForgeX static Mermaid rendering to OfficeIMO Markdown pipelines.</summary>
public static class MermaidMarkdownAdapter {
    /// <summary>Replaces successfully rendered Mermaid fences in place with embedded SVG images.</summary>
    /// <remarks>Nested fences and captions are preserved. Unsupported or invalid fences remain source blocks and report diagnostics.</remarks>
    public static MarkdownDocument Materialize(MarkdownDocument document, Action<MarkupDiagnostic>? diagnostic = null) =>
        Materialize(document, MermaidMarkdownImageFormat.Svg, diagnostic);

    /// <summary>Replaces successfully rendered Mermaid fences in place with embedded images of the requested format.</summary>
    /// <remarks>Nested fences and captions are preserved. Unsupported or invalid fences remain source blocks and report diagnostics.</remarks>
    public static MarkdownDocument Materialize(MarkdownDocument document, MermaidMarkdownImageFormat format, Action<MarkupDiagnostic>? diagnostic = null) {
        ArgumentNullException.ThrowIfNull(document);
        return document.Rewrite(new ImageRewriter(format, diagnostic));
    }

    /// <summary>Creates a reusable reader transform that embeds SVG images, for converters that accept MarkdownLoadOptions.</summary>
    public static IMarkdownDocumentTransform CreateTransform(Action<MarkupDiagnostic>? diagnostic = null) =>
        CreateTransform(MermaidMarkdownImageFormat.Svg, diagnostic);

    /// <summary>Creates a reusable reader transform that embeds images of the requested format.</summary>
    public static IMarkdownDocumentTransform CreateTransform(MermaidMarkdownImageFormat format, Action<MarkupDiagnostic>? diagnostic = null) =>
        new ImageTransform(format, diagnostic);

    /// <summary>Registers static SVG image rendering and disables the Mermaid browser runtime.</summary>
    /// <remarks>Each SVG is an embedded image document, isolating its IDs from other diagrams. Failed renders retain the ordinary source fence.</remarks>
    public static void ConfigureHtml(MarkdownRendererOptions options, Action<MarkupDiagnostic>? diagnostic = null) {
        ArgumentNullException.ThrowIfNull(options);
        options.Mermaid.Enabled = false;
        const string name = "ChartForgeX static Mermaid";
        for (int index = options.FencedCodeBlockRenderers.Count - 1; index >= 0; index--) {
            if (options.FencedCodeBlockRenderers[index].Name == name) options.FencedCodeBlockRenderers.RemoveAt(index);
        }
        options.FencedCodeBlockRenderers.Add(new MarkdownFencedCodeBlockRenderer(name, new[] { "mermaid" }, (match, _) => {
            var artifact = Render(match.InfoString, match.RawContent, match.SourceSpan, match.ContentSourceSpan, diagnostic);
            if (artifact == null) return match.OriginalHtml;
            try {
                string svg = artifact.ToSvg();
                string source = "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
                string alt = System.Net.WebUtility.HtmlEncode(AlternativeText(artifact));
                return "<img class=\"officeimo-mermaid\" style=\"max-width:100%;height:auto\" src=\"" + source + "\" alt=\"" + alt + "\" />";
            } catch (Exception ex) when (IsRenderFailure(ex)) {
                ReportFailure(match.SourceSpan, ex, diagnostic);
                return match.OriginalHtml;
            }
        }));
    }

    private static VisualArtifact? Render(string info, string source, MarkdownSourceSpan? span,
        MarkdownSourceSpan? contentSpan, Action<MarkupDiagnostic>? diagnostic) {
        var fence = MarkdownCodeFenceInfo.Parse(info);
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in fence.Attributes) if (pair.Value != null) attributes[pair.Key] = pair.Value;
        if (fence.ElementId != null) attributes["id"] = fence.ElementId;
        int fenceLine = span?.StartLine ?? 1;
        int firstContentLine = contentSpan?.StartLine ?? fenceLine + 1;
        var block = new VisualMarkupBlock(VisualMarkupKind.Mermaid, "mermaid", info, 0, source,
            fenceLine, firstContentLine, contentSpan?.EndLine ?? firstContentLine + source.Split('\n').Length - 1, attributes);
        var result = new VisualMarkupParseResult();
        new MermaidVisualMarkupBlockParser().Parse(block, result);
        foreach (var item in result.Diagnostics) diagnostic?.Invoke(item);
        return result.HasErrors || result.Artifacts.Count != 1 ? null : result.Artifacts[0];
    }

    private static string AlternativeText(VisualArtifact artifact) =>
        artifact.Accessibility.Description ?? artifact.Accessibility.Name ?? artifact.Title;

    private static bool IsRenderFailure(Exception exception) =>
        exception is ArgumentException || exception is InvalidOperationException || exception is OverflowException;

    private static void ReportFailure(MarkdownSourceSpan? span, Exception exception, Action<MarkupDiagnostic>? diagnostic) =>
        diagnostic?.Invoke(new MarkupDiagnostic {
            Line = span?.StartLine ?? 1,
            Severity = VisualDiagnosticSeverity.Error,
            Message = exception.Message
        });

    private sealed class ImageTransform(MermaidMarkdownImageFormat format, Action<MarkupDiagnostic>? diagnostic) : IMarkdownDocumentTransform {
        public MarkdownDocument Transform(MarkdownDocument document, MarkdownDocumentTransformContext context) =>
            Materialize(document, format, diagnostic);
    }

    private sealed class ImageRewriter(MermaidMarkdownImageFormat format, Action<MarkupDiagnostic>? diagnostic) : MarkdownRewriter {
        protected override IMarkdownBlock RewriteCurrentBlock(IMarkdownBlock block) {
            string language, info, source;
            string? caption;
            MarkdownSourceSpan? span, contentSpan;
            if (block is CodeBlock code) {
                language = code.Language; info = code.InfoString; source = code.Content; caption = code.Caption;
                span = code.SourceSpan; contentSpan = code.ContentSourceSpan;
            } else if (block is SemanticFencedBlock semantic) {
                language = semantic.Language; info = semantic.InfoString; source = semantic.Content; caption = semantic.Caption;
                span = semantic.SourceSpan; contentSpan = semantic.ContentSourceSpan;
            } else return block;
            if (!string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase)) return block;
            var artifact = Render(info, source, span, contentSpan, diagnostic);
            if (artifact == null) return block;
            try {
                string imageSource = format == MermaidMarkdownImageFormat.Svg
                    ? "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(artifact.ToSvg()))
                    : "data:image/png;base64," + Convert.ToBase64String(artifact.ToPng());
                return new ImageBlock(imageSource,
                    AlternativeText(artifact), artifact.Title) { Caption = caption };
            } catch (Exception ex) when (IsRenderFailure(ex)) {
                ReportFailure(span, ex, diagnostic);
                return block;
            }
        }
    }
}
