using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>Executive-summary layout variants.</summary>
    public enum PowerPointExecutiveSummaryLayoutVariant {
        /// <summary>Resolve the layout from content and design intent.</summary>
        Auto,
        /// <summary>Lead with metrics followed by supporting decision points.</summary>
        MetricLead,
        /// <summary>Pair a decision panel with a concise evidence rail.</summary>
        DecisionBrief
    }
    /// <summary>Chart-story layout variants.</summary>
    public enum PowerPointChartStoryLayoutVariant {
        /// <summary>Resolve the layout from content and design intent.</summary>
        Auto,
        /// <summary>Use the editable chart as the dominant visual.</summary>
        ChartHero,
        /// <summary>Pair the chart with a narrative insight rail.</summary>
        InsightRail
    }
    /// <summary>Comparison layout variants.</summary>
    public enum PowerPointComparisonLayoutVariant {
        /// <summary>Resolve the layout from item count and design intent.</summary>
        Auto,
        /// <summary>Render options as parallel narrative cards.</summary>
        SideBySide,
        /// <summary>Render comparison criteria in an editable matrix.</summary>
        DecisionMatrix
    }
    /// <summary>Screenshot-story layout variants.</summary>
    public enum PowerPointScreenshotStoryLayoutVariant {
        /// <summary>Resolve the layout from annotations and design intent.</summary>
        Auto,
        /// <summary>Use a large annotated screenshot as the hero visual.</summary>
        HeroAnnotated,
        /// <summary>Pair the screenshot with a separate narrative rail.</summary>
        SplitNarrative
    }
    /// <summary>Appendix-table layout variants.</summary>
    public enum PowerPointAppendixTableLayoutVariant {
        /// <summary>Resolve the layout from notes and design intent.</summary>
        Auto,
        /// <summary>Give the editable table the full content width.</summary>
        FullWidth,
        /// <summary>Pair the table with a notes and interpretation rail.</summary>
        NotesRail
    }
    /// <summary>Architecture layout variants.</summary>
    public enum PowerPointArchitectureLayoutVariant {
        /// <summary>Resolve the layout from grouping and graph shape.</summary>
        Auto,
        /// <summary>Arrange editable nodes in named horizontal layers.</summary>
        Layered,
        /// <summary>Arrange editable nodes around a central hub.</summary>
        HubSpoke
    }
    /// <summary>Closing-slide layout variants.</summary>
    public enum PowerPointClosingLayoutVariant {
        /// <summary>Resolve the layout from the closing content.</summary>
        Auto,
        /// <summary>Use an expressive full-slide statement.</summary>
        Statement,
        /// <summary>Pair the closing message with an explicit action panel.</summary>
        ActionPanel
    }
    /// <summary>Semantic image placement policy.</summary>
    public enum PowerPointImagePlacement {
        /// <summary>Fill the target box and crop around the focal point.</summary>
        Fill,
        /// <summary>Fit the complete image inside the target box.</summary>
        Fit,
        /// <summary>Stretch the image to the target box.</summary>
        Stretch
    }

    /// <summary>Normalized annotation anchored to a semantic image.</summary>
    public sealed class PowerPointImageAnnotation {
        /// <summary>Creates an annotation at normalized image coordinates from zero to one.</summary>
        public PowerPointImageAnnotation(double x, double y, string label, string? detail = null,
            string? color = null) {
            if (x < 0D || x > 1D) throw new ArgumentOutOfRangeException(nameof(x));
            if (y < 0D || y > 1D) throw new ArgumentOutOfRangeException(nameof(y));
            X = x;
            Y = y;
            Label = string.IsNullOrWhiteSpace(label)
                ? throw new ArgumentException("Annotation label cannot be empty.", nameof(label))
                : label;
            Detail = detail;
            Color = color;
        }

        /// <summary>Normalized horizontal anchor.</summary>
        public double X { get; }
        /// <summary>Normalized vertical anchor.</summary>
        public double Y { get; }
        /// <summary>Short annotation label.</summary>
        public string Label { get; }
        /// <summary>Optional supporting detail.</summary>
        public string? Detail { get; }
        /// <summary>Optional accent color.</summary>
        public string? Color { get; }
    }

    /// <summary>Real image asset with crop, focal-point, caption, provenance, alt text, and annotations.</summary>
    public sealed class PowerPointImageAsset {
        private readonly List<PowerPointImageAnnotation> _annotations = new();

        /// <summary>Creates a semantic image asset.</summary>
        public PowerPointImageAsset(string path, string alternativeText) {
            Path = string.IsNullOrWhiteSpace(path)
                ? throw new ArgumentException("Image path cannot be empty.", nameof(path))
                : path;
            AlternativeText = string.IsNullOrWhiteSpace(alternativeText)
                ? throw new ArgumentException("Alternative text cannot be empty.", nameof(alternativeText))
                : alternativeText;
        }

        /// <summary>Source image path.</summary>
        public string Path { get; }
        /// <summary>Alternative text applied to the native picture.</summary>
        public string AlternativeText { get; }
        /// <summary>Optional visible caption.</summary>
        public string? Caption { get; set; }
        /// <summary>Optional source or provenance line.</summary>
        public string? Provenance { get; set; }
        /// <summary>Image placement policy.</summary>
        public PowerPointImagePlacement Placement { get; set; } = PowerPointImagePlacement.Fill;
        /// <summary>Normalized horizontal focal point.</summary>
        public double FocalX { get; set; } = 0.5D;
        /// <summary>Normalized vertical focal point.</summary>
        public double FocalY { get; set; } = 0.5D;
        /// <summary>Annotations rendered over the image.</summary>
        public IReadOnlyList<PowerPointImageAnnotation> Annotations => _annotations;

        /// <summary>Adds an annotation and returns this asset.</summary>
        public PowerPointImageAsset Annotate(PowerPointImageAnnotation annotation) {
            _annotations.Add(annotation ?? throw new ArgumentNullException(nameof(annotation)));
            return this;
        }

        internal void Validate() {
            if (FocalX < 0D || FocalX > 1D) throw new ArgumentOutOfRangeException(nameof(FocalX));
            if (FocalY < 0D || FocalY > 1D) throw new ArgumentOutOfRangeException(nameof(FocalY));
            if (!File.Exists(Path)) throw new FileNotFoundException("Semantic image asset was not found.", Path);
        }
    }

    /// <summary>Executive-summary content.</summary>
    public sealed class PowerPointExecutiveSummaryContent {
        /// <summary>Creates summary content from metrics and decision points.</summary>
        public PowerPointExecutiveSummaryContent(IEnumerable<PowerPointMetric> metrics,
            IEnumerable<PowerPointCardContent> points, string? lead = null) {
            Metrics = Materialize(metrics, nameof(metrics));
            Points = Materialize(points, nameof(points));
            Lead = lead;
            if (Metrics.Count == 0 && Points.Count == 0 && string.IsNullOrWhiteSpace(Lead)) {
                throw new ArgumentException(
                    "Executive-summary content requires a lead, metric, or decision point.", nameof(points));
            }
        }
        /// <summary>Top-line metrics.</summary>
        public IReadOnlyList<PowerPointMetric> Metrics { get; }
        /// <summary>Decision, status, or recommendation points.</summary>
        public IReadOnlyList<PowerPointCardContent> Points { get; }
        /// <summary>Optional lead statement.</summary>
        public string? Lead { get; }

        private static IReadOnlyList<T> Materialize<T>(IEnumerable<T> source, string name) {
            if (source == null) throw new ArgumentNullException(name);
            return new ReadOnlyCollection<T>(source.Where(item => item != null).ToList());
        }
    }

    /// <summary>Editable chart plus its narrative and accessibility context.</summary>
    public sealed class PowerPointChartStoryContent {
        /// <summary>Creates a chart story.</summary>
        public PowerPointChartStoryContent(OfficeChartKind chartKind, PowerPointChartData data,
            IEnumerable<string>? insights = null) {
            ChartKind = chartKind;
            Data = data ?? throw new ArgumentNullException(nameof(data));
            Insights = new ReadOnlyCollection<string>((insights ?? Array.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value)).ToList());
        }
        /// <summary>Shared chart family.</summary>
        public OfficeChartKind ChartKind { get; }
        /// <summary>Native chart data.</summary>
        public PowerPointChartData Data { get; }
        /// <summary>Narrative insights displayed beside or below the chart.</summary>
        public IReadOnlyList<string> Insights { get; }
        /// <summary>Optional visible caption.</summary>
        public string? Caption { get; set; }
        /// <summary>Optional source/provenance line.</summary>
        public string? Provenance { get; set; }
        /// <summary>Alternative text applied to the native chart.</summary>
        public string? AlternativeText { get; set; }
        /// <summary>Plain-language summary of the represented data.</summary>
        public string? DataSummary { get; set; }
    }

    /// <summary>One option in a semantic comparison.</summary>
    public sealed class PowerPointComparisonItem {
        /// <summary>Creates a comparison item.</summary>
        public PowerPointComparisonItem(string title, string? summary = null,
            IEnumerable<string>? strengths = null, IEnumerable<string>? tradeoffs = null) {
            Title = string.IsNullOrWhiteSpace(title)
                ? throw new ArgumentException("Comparison title cannot be empty.", nameof(title))
                : title;
            Summary = summary;
            Strengths = Clean(strengths);
            Tradeoffs = Clean(tradeoffs);
        }
        /// <summary>Option title.</summary>
        public string Title { get; }
        /// <summary>Option summary.</summary>
        public string? Summary { get; }
        /// <summary>Positive evidence.</summary>
        public IReadOnlyList<string> Strengths { get; }
        /// <summary>Tradeoffs or risks.</summary>
        public IReadOnlyList<string> Tradeoffs { get; }
        private static IReadOnlyList<string> Clean(IEnumerable<string>? values) =>
            new ReadOnlyCollection<string>((values ?? Array.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value)).ToList());
    }

    /// <summary>Native appendix table data independent of CLR row types.</summary>
    public sealed class PowerPointTableData {
        /// <summary>Creates table data from headers and string rows.</summary>
        public PowerPointTableData(IEnumerable<string> headers, IEnumerable<IEnumerable<string>> rows) {
            Headers = new ReadOnlyCollection<string>((headers ?? throw new ArgumentNullException(nameof(headers)))
                .ToList());
            if (Headers.Count == 0) throw new ArgumentException("At least one table header is required.", nameof(headers));
            var materializedRows = new List<IReadOnlyList<string>>();
            foreach (IEnumerable<string> row in rows ?? throw new ArgumentNullException(nameof(rows))) {
                List<string> values = row.ToList();
                if (values.Count != Headers.Count) {
                    throw new ArgumentException("Every table row must match the header count.", nameof(rows));
                }
                materializedRows.Add(new ReadOnlyCollection<string>(values));
            }
            Rows = new ReadOnlyCollection<IReadOnlyList<string>>(materializedRows);
        }
        /// <summary>Column headers.</summary>
        public IReadOnlyList<string> Headers { get; }
        /// <summary>Table rows.</summary>
        public IReadOnlyList<IReadOnlyList<string>> Rows { get; }
        /// <summary>Optional caption.</summary>
        public string? Caption { get; set; }
        /// <summary>Optional source/provenance line.</summary>
        public string? Provenance { get; set; }
        /// <summary>Optional notes displayed beside the table.</summary>
        public IReadOnlyList<string> Notes { get; set; } = Array.Empty<string>();
    }

    /// <summary>One editable architecture node.</summary>
    public sealed class PowerPointArchitectureNode {
        /// <summary>Creates an architecture node.</summary>
        public PowerPointArchitectureNode(string id, string title, string? body = null, string? group = null) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Node id cannot be empty.", nameof(id)) : id;
            Title = string.IsNullOrWhiteSpace(title) ? throw new ArgumentException("Node title cannot be empty.", nameof(title)) : title;
            Body = body;
            Group = group;
        }
        /// <summary>Stable node id used by edges.</summary>
        public string Id { get; }
        /// <summary>Node title.</summary>
        public string Title { get; }
        /// <summary>Optional node body.</summary>
        public string? Body { get; }
        /// <summary>Optional layer or group.</summary>
        public string? Group { get; }
    }

    /// <summary>One directed relationship between architecture nodes.</summary>
    public sealed class PowerPointArchitectureEdge {
        /// <summary>Creates an edge.</summary>
        public PowerPointArchitectureEdge(string fromId, string toId, string? label = null) {
            FromId = fromId ?? throw new ArgumentNullException(nameof(fromId));
            ToId = toId ?? throw new ArgumentNullException(nameof(toId));
            Label = label;
        }
        /// <summary>Source node id.</summary>
        public string FromId { get; }
        /// <summary>Target node id.</summary>
        public string ToId { get; }
        /// <summary>Optional relationship label.</summary>
        public string? Label { get; }
    }

    /// <summary>Semantic architecture graph.</summary>
    public sealed class PowerPointArchitectureContent {
        /// <summary>Creates an architecture graph.</summary>
        public PowerPointArchitectureContent(IEnumerable<PowerPointArchitectureNode> nodes,
            IEnumerable<PowerPointArchitectureEdge>? edges = null) {
            Nodes = new ReadOnlyCollection<PowerPointArchitectureNode>((nodes ?? throw new ArgumentNullException(nameof(nodes))).ToList());
            if (Nodes.Count == 0) throw new ArgumentException("Architecture requires at least one node.", nameof(nodes));
            Edges = new ReadOnlyCollection<PowerPointArchitectureEdge>((edges ?? Array.Empty<PowerPointArchitectureEdge>()).ToList());
            var ids = new HashSet<string>(Nodes.Select(node => node.Id), StringComparer.OrdinalIgnoreCase);
            foreach (PowerPointArchitectureEdge edge in Edges) {
                if (!ids.Contains(edge.FromId) || !ids.Contains(edge.ToId))
                    throw new ArgumentException("Architecture edge references an unknown node.", nameof(edges));
            }
        }
        /// <summary>Architecture nodes.</summary>
        public IReadOnlyList<PowerPointArchitectureNode> Nodes { get; }
        /// <summary>Architecture relationships.</summary>
        public IReadOnlyList<PowerPointArchitectureEdge> Edges { get; }
    }

    /// <summary>Closing-slide message and call to action.</summary>
    public sealed class PowerPointClosingContent {
        /// <summary>Creates closing content.</summary>
        public PowerPointClosingContent(string statement, string? callToAction = null, string? contact = null) {
            Statement = string.IsNullOrWhiteSpace(statement)
                ? throw new ArgumentException("Closing statement cannot be empty.", nameof(statement))
                : statement;
            CallToAction = callToAction;
            Contact = contact;
        }
        /// <summary>Primary closing statement.</summary>
        public string Statement { get; }
        /// <summary>Optional explicit next action.</summary>
        public string? CallToAction { get; }
        /// <summary>Optional contact or handoff line.</summary>
        public string? Contact { get; }
    }

    /// <summary>Executive-summary options.</summary>
    public sealed class PowerPointExecutiveSummarySlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>Gets or sets the layout variant.</summary>
        public PowerPointExecutiveSummaryLayoutVariant Variant { get; set; } = PowerPointExecutiveSummaryLayoutVariant.Auto;
    }
    /// <summary>Chart-story options.</summary>
    public sealed class PowerPointChartStorySlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>Gets or sets the layout variant.</summary>
        public PowerPointChartStoryLayoutVariant Variant { get; set; } = PowerPointChartStoryLayoutVariant.Auto;
    }
    /// <summary>Comparison options.</summary>
    public sealed class PowerPointComparisonSlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>Gets or sets the layout variant.</summary>
        public PowerPointComparisonLayoutVariant Variant { get; set; } = PowerPointComparisonLayoutVariant.Auto;
    }
    /// <summary>Screenshot-story options.</summary>
    public sealed class PowerPointScreenshotStorySlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>Gets or sets the layout variant.</summary>
        public PowerPointScreenshotStoryLayoutVariant Variant { get; set; } = PowerPointScreenshotStoryLayoutVariant.Auto;
    }
    /// <summary>Appendix-table options.</summary>
    public sealed class PowerPointAppendixTableSlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>Gets or sets the layout variant.</summary>
        public PowerPointAppendixTableLayoutVariant Variant { get; set; } = PowerPointAppendixTableLayoutVariant.Auto;
    }
    /// <summary>Architecture options.</summary>
    public sealed class PowerPointArchitectureSlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>Gets or sets the layout variant.</summary>
        public PowerPointArchitectureLayoutVariant Variant { get; set; } = PowerPointArchitectureLayoutVariant.Auto;
    }
    /// <summary>Closing-slide options.</summary>
    public sealed class PowerPointClosingSlideOptions : PowerPointDesignerSlideOptions {
        /// <summary>Gets or sets the layout variant.</summary>
        public PowerPointClosingLayoutVariant Variant { get; set; } = PowerPointClosingLayoutVariant.Auto;
    }
}
