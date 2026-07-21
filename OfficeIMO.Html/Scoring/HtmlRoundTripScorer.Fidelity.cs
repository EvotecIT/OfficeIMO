using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

public static partial class HtmlRoundTripScorer {
    private static readonly string[] FidelityStyleProperties = {
        "color", "background-color", "font-family", "font-size", "font-style", "font-weight",
        "text-decoration-line", "text-align", "vertical-align", "white-space", "direction",
        "border-color", "border-style", "border-width"
    };

    private static readonly string[] GeometryAttributes = {
        "data-officeimo-left", "data-officeimo-top", "data-officeimo-width", "data-officeimo-height",
        "data-officeimo-row", "data-officeimo-column", "data-officeimo-rotation",
        "data-officeimo-flip-horizontal", "data-officeimo-flip-vertical",
        "data-officeimo-layer-index", "data-officeimo-layer-kind"
    };

    private static HtmlRoundTripScore ApplyFidelityV2(
        HtmlRoundTripScore structuralScore,
        HtmlConversionDocument sourceDocument,
        HtmlConversionDocument targetDocument,
        HtmlArtifactReloadEvidence? artifactReloadEvidence) {
        Dictionary<string, double> metrics = structuralScore.Metrics
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
        Dictionary<string, double> dimensions = structuralScore.Dimensions
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);

        AddFidelityDimension(metrics, dimensions, "styles",
            ExtractStyleSignatures(targetDocument.SemanticDocument),
            ExtractStyleSignatures(sourceDocument.SemanticDocument));
        AddFidelityDimension(metrics, dimensions, "resources",
            ExtractResourceSignatures(targetDocument),
            ExtractResourceSignatures(sourceDocument));

        IHtmlDocument sourceDom = sourceDocument.CreatePolicyNormalizedDocumentForConversion();
        IHtmlDocument targetDom = targetDocument.CreatePolicyNormalizedDocumentForConversion();
        AddFidelityDimension(metrics, dimensions, "annotations",
            ExtractAnnotationSignatures(targetDom),
            ExtractAnnotationSignatures(sourceDom));
        AddFidelityDimension(metrics, dimensions, "formulas",
            ExtractFormulaSignatures(targetDom),
            ExtractFormulaSignatures(sourceDom));
        AddFidelityDimension(metrics, dimensions, "charts",
            ExtractChartSignatures(targetDom),
            ExtractChartSignatures(sourceDom));
        AddFidelityDimension(metrics, dimensions, "geometry",
            ExtractGeometrySignatures(targetDom),
            ExtractGeometrySignatures(sourceDom));

        bool reloadVerified = artifactReloadEvidence?.ReloadSucceeded == true
            && !string.IsNullOrWhiteSpace(artifactReloadEvidence.ReloadedHtml);
        string? artifactKind = artifactReloadEvidence?.ArtifactKind;
        if (artifactReloadEvidence != null) {
            double reloadScore = 0D;
            if (artifactReloadEvidence.ReloadSucceeded
                && !string.IsNullOrWhiteSpace(artifactReloadEvidence.ReloadedHtml)) {
                HtmlConversionDocument reloaded = HtmlConversionDocument.Parse(artifactReloadEvidence.ReloadedHtml!);
                reloadScore = Compare(sourceDocument, reloaded).Score;
            }

            AddMetric(metrics, "artifact-reload", reloadScore);
            dimensions["artifact-reload"] = reloadScore;
        }

        int compared = metrics.Count;
        int matched = metrics.Values.Count(value => value >= 0.95D);
        double score = dimensions.Count == 0 ? 1D : dimensions.Values.Average();
        return new HtmlRoundTripScore(
            score,
            structuralScore.SourceNodeCount,
            structuralScore.TargetNodeCount,
            matched,
            compared,
            metrics,
            dimensions,
            reloadVerified,
            artifactKind);
    }

    private static void AddFidelityDimension(
        IDictionary<string, double> metrics,
        IDictionary<string, double> dimensions,
        string name,
        IReadOnlyList<string> actual,
        IReadOnlyList<string> expected) {
        if (actual.Count == 0 && expected.Count == 0) return;
        double similarity = SignatureSimilarity(actual, expected);
        AddMetric(metrics, name, similarity);
        dimensions[name] = similarity;
    }

    private static IReadOnlyList<string> ExtractStyleSignatures(HtmlSemanticDocument document) {
        var signatures = new List<string>();
        var visited = new HashSet<IElement>();
        foreach (HtmlSemanticSection section in document.Sections) {
            foreach (HtmlSemanticBlock block in section.Blocks) {
                AppendStyleSignatures(block, signatures, visited);
            }
        }
        foreach (HtmlSemanticBlock table in document.RootTables) {
            AppendStyleSignatures(table, signatures, visited);
        }
        return signatures;
    }

    private static void AppendStyleSignatures(
        HtmlSemanticBlock block,
        ICollection<string> signatures,
        ISet<IElement> visited) {
        if (!visited.Add(block.SourceElement)) return;
        AddStyleSignature(signatures, "block:" + block.Kind, block.Style, string.Empty);
        foreach (HtmlSemanticRun run in block.Runs) {
            string flags = (run.Bold ? "b" : string.Empty)
                + (run.Italic ? "i" : string.Empty)
                + (run.Underline ? "u" : string.Empty)
                + (run.Strikethrough ? "s" : string.Empty)
                + (run.Superscript ? "sup" : string.Empty)
                + (run.Subscript ? "sub" : string.Empty);
            AddStyleSignature(signatures, "run", run.Style, flags);
        }
        if (block.Table != null) {
            foreach (HtmlSemanticTableRow row in block.Table.Rows) {
                foreach (HtmlSemanticTableCell cell in row.Cells) {
                    AddStyleSignature(signatures, cell.IsHeader ? "cell:header" : "cell:data", cell.Style, string.Empty);
                    foreach (HtmlSemanticRun run in cell.Runs) {
                        string flags = (run.Bold ? "b" : string.Empty)
                            + (run.Italic ? "i" : string.Empty)
                            + (run.Underline ? "u" : string.Empty)
                            + (run.Strikethrough ? "s" : string.Empty)
                            + (run.Superscript ? "sup" : string.Empty)
                            + (run.Subscript ? "sub" : string.Empty);
                        AddStyleSignature(signatures, "cell:run", run.Style, flags);
                    }
                }
            }
        }
        foreach (HtmlSemanticBlock child in block.Children) {
            AppendStyleSignatures(child, signatures, visited);
        }
    }

    private static void AddStyleSignature(
        ICollection<string> signatures,
        string owner,
        HtmlComputedStyle? style,
        string flags) {
        var parts = new List<string>();
        foreach (string property in FidelityStyleProperties) {
            string value = style?.GetValue(property) ?? string.Empty;
            if (value.Length > 0) parts.Add(property + "=" + value.Trim().ToLowerInvariant());
        }
        if (flags.Length > 0) parts.Add("format=" + flags);
        if (parts.Count > 0) signatures.Add(owner + "|" + string.Join("|", parts));
    }

    private static IReadOnlyList<string> ExtractResourceSignatures(HtmlConversionDocument document) {
        var signatures = new List<string>();
        foreach (HtmlResourceReference resource in document.ResourceManifest.Resources) {
            if (resource.Kind == HtmlResourceKind.Hyperlink) continue;
            string source = resource.IsAllowed && resource.ResolvedSource.Length > 0
                ? resource.ResolvedSource
                : resource.Source;
            signatures.Add("reference|" + resource.Kind + "|" + resource.ElementName.ToLowerInvariant()
                + "|" + resource.AttributeName.ToLowerInvariant() + "|" + NormalizeResourceIdentity(source)
                + "|" + (resource.IsAllowed ? "allowed" : "blocked"));
        }
        foreach (HtmlSemanticResource resource in document.SemanticDocument.Resources) {
            signatures.Add("semantic|" + resource.Kind + "|" + NormalizeResourceIdentity(resource.Source)
                + "|" + resource.MediaType.Trim().ToLowerInvariant()
                + "|" + NormalizeText(resource.AlternateText));
        }
        return signatures;
    }

    private static string NormalizeResourceIdentity(string source) {
        string normalized = (source ?? string.Empty).Trim();
        if (!normalized.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return normalized;
        int comma = normalized.IndexOf(',');
        string media = comma > 5 ? normalized.Substring(5, comma - 5).ToLowerInvariant() : string.Empty;
        return "data:" + media + "#sha256=" + Hash(normalized);
    }

    private static IReadOnlyList<string> ExtractAnnotationSignatures(IHtmlDocument document) {
        var signatures = new List<string>();
        foreach (IElement element in HtmlSemanticFeatureLocator.FindAnnotations(document)
            .Concat(HtmlSemanticFeatureLocator.FindComments(document)).Distinct()) {
            string name = element.TagName.ToLowerInvariant();
            string kind = name == "ins" || name == "del" ? name
                : element.HasAttribute("data-officeimo-comment") || IsInsideClass(element, "officeimo-comments") ? "comment"
                : element.HasAttribute("data-officeimo-bookmark") ? "bookmark"
                : "annotation";
            signatures.Add(kind + "|" + ReadIdentity(element) + "|" + NormalizeText(element.TextContent));
        }
        return signatures;
    }

    private static IReadOnlyList<string> ExtractFormulaSignatures(IHtmlDocument document) {
        var signatures = new List<string>();
        foreach (IElement element in HtmlSemanticFeatureLocator.FindFormulas(document)) {
            bool inventory = IsInsideClass(element, "officeimo-formulas");
            signatures.Add((inventory ? "inventory|" : "cell|") + ReadIdentity(element) + "|"
                + (element.GetAttribute("data-officeimo-formula")
                    ?? element.GetAttribute("data-officeimo-value")
                    ?? NormalizeText(element.QuerySelector("code")?.TextContent ?? element.TextContent)));
        }
        return signatures;
    }

    private static IReadOnlyList<string> ExtractChartSignatures(IHtmlDocument document) {
        var signatures = new List<string>();
        foreach (IElement element in HtmlSemanticFeatureLocator.FindCharts(document)) {
            signatures.Add(CreateChartSignature(element));
        }
        return signatures;
    }

    private static string CreateChartSignature(IElement element) =>
        "chart|" + (element.GetAttribute("data-officeimo-chart-type")
            ?? element.GetAttribute("data-officeimo-chart-kind")
            ?? string.Empty)
        + "|" + ReadIdentity(element)
        + "|" + string.Join("|", GeometryAttributes.Select(attribute =>
            attribute + "=" + NormalizeNumericAttribute(element.GetAttribute(attribute))))
        + "|" + NormalizeText(element.TextContent);

    private static IReadOnlyList<string> ExtractGeometrySignatures(IHtmlDocument document) {
        var signatures = new List<string>();
        foreach (IElement element in document.QuerySelectorAll("*")) {
            if (!GeometryAttributes.Any(element.HasAttribute)) continue;
            var parts = new List<string>();
            foreach (string attribute in GeometryAttributes) {
                if (!element.HasAttribute(attribute)) continue;
                parts.Add(attribute + "=" + NormalizeNumericAttribute(element.GetAttribute(attribute)));
            }
            signatures.Add(element.TagName.ToLowerInvariant() + "|" + ReadIdentity(element) + "|" + string.Join("|", parts));
        }
        return signatures;
    }

    private static string NormalizeNumericAttribute(string? value) {
        string text = (value ?? string.Empty).Trim();
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            ? number.ToString("G17", CultureInfo.InvariantCulture)
            : text.ToLowerInvariant();
    }

    private static string ReadIdentity(IElement element) =>
        element.GetAttribute("data-officeimo-cell")
        ?? element.GetAttribute("id")
        ?? element.GetAttribute("data-officeimo-layer-kind")
        ?? element.TagName.ToLowerInvariant();

    private static bool IsInsideClass(IElement element, string className) {
        for (IElement? current = element.ParentElement; current != null; current = current.ParentElement) {
            if (current.ClassList.Contains(className)) return true;
        }
        return false;
    }
}
