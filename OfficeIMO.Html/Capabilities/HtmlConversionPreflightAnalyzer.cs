using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal static class HtmlConversionPreflightAnalyzer {
    private const string ComponentName = "OfficeIMO.Html.Preflight";

    internal static HtmlConversionPreflight Analyze(HtmlConversionDocument document, HtmlConversionTarget target) {
        HtmlTargetCapabilityContract contract = HtmlTargetCapabilityContracts.Get(target);
        HtmlSemanticDocument semantic = document.SemanticDocument;
        IHtmlDocument dom = document.CreatePolicyNormalizedDocumentForConversion();
        IReadOnlyList<HtmlSemanticBlock> blocks = Flatten(semantic.Sections.SelectMany(section => section.Blocks)).ToList();
        var features = new List<HtmlFeaturePreflightResult>();
        var diagnostics = new List<HtmlDiagnostic>();
        foreach (HtmlSemanticFeature feature in (HtmlSemanticFeature[])Enum.GetValues(typeof(HtmlSemanticFeature))) {
            FeatureEvidence evidence = Count(feature, document, dom, semantic, blocks);
            HtmlConversionPreflightOutcome outcome = Map(contract.GetSupport(feature));
            var result = new HtmlFeaturePreflightResult(feature, evidence.Count > 0, evidence.Count, outcome, evidence.Location);
            features.Add(result);
            if (evidence.Count == 0 || outcome == HtmlConversionPreflightOutcome.Supported) continue;

            bool omitted = outcome == HtmlConversionPreflightOutcome.Omitted;
            diagnostics.Add(new HtmlDiagnostic(
                ComponentName,
                omitted ? HtmlConversionDiagnosticCodes.ContentOmitted : HtmlConversionDiagnosticCodes.ContentApproximated,
                feature + " is predicted to be " + (omitted ? "omitted" : "approximated") + " by the " + target + " target.",
                HtmlDiagnosticSeverity.Warning,
                source: evidence.Location?.Selector,
                detail: "target=" + target + "; feature=" + feature + "; occurrences=" + evidence.Count,
                lossKind: omitted ? HtmlConversionLossKind.Omission : HtmlConversionLossKind.Approximation,
                sourceLocation: evidence.Location,
                targetAddress: "preflight:" + target.ToString().ToLowerInvariant()));
        }
        return new HtmlConversionPreflight(target, contract, features.AsReadOnly(), diagnostics.AsReadOnly());
    }

    private static FeatureEvidence Count(HtmlSemanticFeature feature, HtmlConversionDocument document,
        IHtmlDocument dom, HtmlSemanticDocument semantic, IReadOnlyList<HtmlSemanticBlock> blocks) {
        switch (feature) {
            case HtmlSemanticFeature.Metadata:
                return Evidence(semantic.Metadata.Count, blocks.FirstOrDefault());
            case HtmlSemanticFeature.Sections:
                return new FeatureEvidence(semantic.Sections.Count, semantic.Sections.FirstOrDefault()?.SourceLocation);
            case HtmlSemanticFeature.Headings:
                return Evidence(blocks.Count(block => block.Kind == HtmlSemanticBlockKind.Heading), blocks.FirstOrDefault(block => block.Kind == HtmlSemanticBlockKind.Heading));
            case HtmlSemanticFeature.Paragraphs:
                return Evidence(blocks.Count(block => block.Kind == HtmlSemanticBlockKind.Paragraph), blocks.FirstOrDefault(block => block.Kind == HtmlSemanticBlockKind.Paragraph));
            case HtmlSemanticFeature.RichText:
                return RunEvidence(blocks, run => run.Bold || run.Italic || run.Underline || run.Strikethrough || run.Superscript || run.Subscript);
            case HtmlSemanticFeature.Links:
                return RunEvidence(blocks, run => !string.IsNullOrWhiteSpace(run.Hyperlink));
            case HtmlSemanticFeature.Lists:
                return Evidence(blocks.Count(block => block.Kind == HtmlSemanticBlockKind.List), blocks.FirstOrDefault(block => block.Kind == HtmlSemanticBlockKind.List));
            case HtmlSemanticFeature.Tables:
                return Evidence(semantic.RootTables.Count, semantic.RootTables.FirstOrDefault());
            case HtmlSemanticFeature.Images:
                return ResourceEvidence(semantic, HtmlResourceKind.Image);
            case HtmlSemanticFeature.Media:
                return ResourceEvidence(semantic, HtmlResourceKind.Media);
            case HtmlSemanticFeature.Forms:
                return Evidence(blocks.Count(block => block.Kind == HtmlSemanticBlockKind.Form), blocks.FirstOrDefault(block => block.Kind == HtmlSemanticBlockKind.Form));
            case HtmlSemanticFeature.Notes:
                return Evidence(blocks.Count(block => block.Kind == HtmlSemanticBlockKind.Note), blocks.FirstOrDefault(block => block.Kind == HtmlSemanticBlockKind.Note));
            case HtmlSemanticFeature.Comments:
                return ElementEvidence(HtmlSemanticFeatureLocator.FindComments(dom));
            case HtmlSemanticFeature.Annotations:
                return ElementEvidence(HtmlSemanticFeatureLocator.FindAnnotations(dom));
            case HtmlSemanticFeature.Formulas:
                return ElementEvidence(HtmlSemanticFeatureLocator.FindFormulas(dom));
            case HtmlSemanticFeature.Charts:
                return ElementEvidence(HtmlSemanticFeatureLocator.FindCharts(dom));
            case HtmlSemanticFeature.Geometry:
                return GeometryEvidence(blocks);
            case HtmlSemanticFeature.Css:
                return Evidence(document.StyleSummary.StyledElementCount, blocks.FirstOrDefault(block => block.Style?.Properties.Count > 0));
            case HtmlSemanticFeature.Resources:
                return new FeatureEvidence(document.ResourceManifest.Resources.Count, semantic.ResourceOccurrences.FirstOrDefault()?.SourceLocation);
            case HtmlSemanticFeature.PagedLayout:
                return ElementEvidence(HtmlSemanticFeatureLocator.FindPagedLayout(dom));
            default:
                return default;
        }
    }

    private static FeatureEvidence GeometryEvidence(IReadOnlyList<HtmlSemanticBlock> blocks) {
        HtmlSemanticBlock? first = blocks.FirstOrDefault(block => HasGeometry(block));
        return new FeatureEvidence(blocks.Count(HasGeometry), first?.SourceLocation);
    }

    private static bool HasGeometry(HtmlSemanticBlock block) {
        if (block.Style == null) return false;
        string position = block.Style.GetValue("position");
        return (!string.IsNullOrWhiteSpace(position) && !string.Equals(position, "static", StringComparison.OrdinalIgnoreCase))
            || !string.IsNullOrWhiteSpace(block.Style.GetValue("width"))
            || !string.IsNullOrWhiteSpace(block.Style.GetValue("height"))
            || !string.IsNullOrWhiteSpace(block.Style.GetValue("transform"));
    }

    private static FeatureEvidence RunEvidence(IReadOnlyList<HtmlSemanticBlock> blocks, Func<HtmlSemanticRun, bool> predicate) {
        int count = 0;
        HtmlSemanticSourceLocation? location = null;
        foreach (HtmlSemanticBlock block in blocks) {
            foreach (HtmlSemanticRun run in block.Runs) {
                if (!predicate(run)) continue;
                count++;
                location ??= run.SourceLocation ?? block.SourceLocation;
            }
        }
        return new FeatureEvidence(count, location);
    }

    private static FeatureEvidence ResourceEvidence(HtmlSemanticDocument document, HtmlResourceKind kind) {
        IReadOnlyList<HtmlSemanticResource> resources = document.ResourceOccurrences.Where(resource => resource.Kind == kind).ToList();
        return new FeatureEvidence(resources.Count, resources.FirstOrDefault()?.SourceLocation);
    }

    private static FeatureEvidence ElementEvidence(IReadOnlyList<IElement> elements) =>
        new FeatureEvidence(elements.Count,
            elements.Count == 0 ? null : HtmlSemanticSourceLocation.FromElement(elements[0]));

    private static FeatureEvidence Evidence(int count, HtmlSemanticBlock? block) => new FeatureEvidence(count, block?.SourceLocation);

    private static IEnumerable<HtmlSemanticBlock> Flatten(IEnumerable<HtmlSemanticBlock> blocks) {
        foreach (HtmlSemanticBlock block in blocks) {
            yield return block;
            foreach (HtmlSemanticBlock child in Flatten(block.Children)) yield return child;
        }
    }

    private static HtmlConversionPreflightOutcome Map(HtmlCapabilitySupportLevel support) => support switch {
        HtmlCapabilitySupportLevel.Supported => HtmlConversionPreflightOutcome.Supported,
        HtmlCapabilitySupportLevel.Approximated => HtmlConversionPreflightOutcome.Approximated,
        _ => HtmlConversionPreflightOutcome.Omitted
    };

    private readonly struct FeatureEvidence {
        internal FeatureEvidence(int count, HtmlSemanticSourceLocation? location) {
            Count = count;
            Location = location;
        }
        internal int Count { get; }
        internal HtmlSemanticSourceLocation? Location { get; }
    }
}
