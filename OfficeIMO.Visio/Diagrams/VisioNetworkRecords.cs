using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Simple data record used to import network zones into <see cref="VisioNetworkDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioNetworkZoneRecord {
        /// <summary>Initializes a network zone import record.</summary>
        public VisioNetworkZoneRecord(string id, string text, int column, int row, int columnSpan, int rowSpan) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Zone id cannot be null or whitespace.", nameof(id)) : id;
            Text = text ?? string.Empty;
            Column = column;
            Row = row;
            ColumnSpan = columnSpan;
            RowSpan = rowSpan;
        }

        /// <summary>Stable zone id.</summary>
        public string Id { get; }

        /// <summary>Visible zone label.</summary>
        public string Text { get; }

        /// <summary>Zero-based grid column.</summary>
        public int Column { get; }

        /// <summary>Zero-based grid row.</summary>
        public int Row { get; }

        /// <summary>Number of grid columns spanned by this zone.</summary>
        public int ColumnSpan { get; }

        /// <summary>Number of grid rows spanned by this zone.</summary>
        public int RowSpan { get; }

        /// <summary>Shape Data rows to apply to the generated zone background.</summary>
        public IDictionary<string, string?> ShapeData { get; } = new Dictionary<string, string?>(StringComparer.Ordinal);

        /// <summary>Optional hyperlink address attached to the generated zone background.</summary>
        public string? HyperlinkAddress { get; set; }

        /// <summary>Optional hyperlink description attached to the generated zone background.</summary>
        public string? HyperlinkDescription { get; set; }
    }

    /// <summary>
    /// Simple data record used to import network nodes into <see cref="VisioNetworkDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioNetworkNodeRecord {
        /// <summary>Initializes a network node import record.</summary>
        public VisioNetworkNodeRecord(string id, string text, int column, int row, VisioNetworkNodeKind kind = VisioNetworkNodeKind.Server) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Node id cannot be null or whitespace.", nameof(id)) : id;
            Text = text ?? string.Empty;
            Column = column;
            Row = row;
            Kind = kind;
        }

        /// <summary>Stable node id.</summary>
        public string Id { get; }

        /// <summary>Visible node label.</summary>
        public string Text { get; }

        /// <summary>Zero-based grid column.</summary>
        public int Column { get; }

        /// <summary>Zero-based grid row.</summary>
        public int Row { get; }

        /// <summary>Semantic node kind.</summary>
        public VisioNetworkNodeKind Kind { get; }

        /// <summary>Shape Data rows to apply to the generated node.</summary>
        public IDictionary<string, string?> ShapeData { get; } = new Dictionary<string, string?>(StringComparer.Ordinal);

        /// <summary>Optional hyperlink address attached to the generated node.</summary>
        public string? HyperlinkAddress { get; set; }

        /// <summary>Optional hyperlink description attached to the generated node.</summary>
        public string? HyperlinkDescription { get; set; }
    }

    /// <summary>
    /// Simple data record used to import network links into <see cref="VisioNetworkDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioNetworkLinkRecord {
        /// <summary>Initializes a network link import record.</summary>
        public VisioNetworkLinkRecord(string id, string fromId, string toId, VisioNetworkLinkKind kind = VisioNetworkLinkKind.Ethernet, string? label = null) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Link id cannot be null or whitespace.", nameof(id)) : id;
            FromId = string.IsNullOrWhiteSpace(fromId) ? throw new ArgumentException("Source node id cannot be null or whitespace.", nameof(fromId)) : fromId;
            ToId = string.IsNullOrWhiteSpace(toId) ? throw new ArgumentException("Target node id cannot be null or whitespace.", nameof(toId)) : toId;
            Kind = kind;
            Label = label;
        }

        /// <summary>Stable link id.</summary>
        public string Id { get; }

        /// <summary>Source node id.</summary>
        public string FromId { get; }

        /// <summary>Target node id.</summary>
        public string ToId { get; }

        /// <summary>Semantic link kind.</summary>
        public VisioNetworkLinkKind Kind { get; }

        /// <summary>Visible link label.</summary>
        public string? Label { get; }

        /// <summary>Shape Data rows to apply to the generated connector.</summary>
        public IDictionary<string, string?> ShapeData { get; } = new Dictionary<string, string?>(StringComparer.Ordinal);

        /// <summary>Optional hyperlink address attached to the generated connector.</summary>
        public string? HyperlinkAddress { get; set; }

        /// <summary>Optional hyperlink description attached to the generated connector.</summary>
        public string? HyperlinkDescription { get; set; }
    }

    /// <summary>
    /// Simple data record used to import network callouts into <see cref="VisioNetworkDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioNetworkCalloutRecord {
        /// <summary>Initializes a placement-based network callout import record.</summary>
        public VisioNetworkCalloutRecord(string id, string targetId, string text, VisioSide placement, double gap = 0.35D) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Callout id cannot be null or whitespace.", nameof(id)) : id;
            TargetId = string.IsNullOrWhiteSpace(targetId) ? throw new ArgumentException("Target node id cannot be null or whitespace.", nameof(targetId)) : targetId;
            Text = text ?? string.Empty;
            Placement = placement;
            Gap = gap;
            UsePlacement = true;
        }

        /// <summary>Initializes an absolute-position network callout import record.</summary>
        public VisioNetworkCalloutRecord(string id, string targetId, string text, double pinX, double pinY) {
            Id = string.IsNullOrWhiteSpace(id) ? throw new ArgumentException("Callout id cannot be null or whitespace.", nameof(id)) : id;
            TargetId = string.IsNullOrWhiteSpace(targetId) ? throw new ArgumentException("Target node id cannot be null or whitespace.", nameof(targetId)) : targetId;
            Text = text ?? string.Empty;
            PinX = pinX;
            PinY = pinY;
        }

        /// <summary>Stable callout id.</summary>
        public string Id { get; }

        /// <summary>Target node id.</summary>
        public string TargetId { get; }

        /// <summary>Visible callout text.</summary>
        public string Text { get; }

        /// <summary>Absolute callout pin X, when <see cref="UsePlacement"/> is false.</summary>
        public double PinX { get; }

        /// <summary>Absolute callout pin Y, when <see cref="UsePlacement"/> is false.</summary>
        public double PinY { get; }

        /// <summary>Requested callout placement side, when <see cref="UsePlacement"/> is true.</summary>
        public VisioSide Placement { get; }

        /// <summary>Requested callout gap from the target node, when <see cref="UsePlacement"/> is true.</summary>
        public double Gap { get; }

        /// <summary>Whether the callout should be placed relative to the target node.</summary>
        public bool UsePlacement { get; }

        /// <summary>Optional explicit callout width.</summary>
        public double? Width { get; set; }

        /// <summary>Optional explicit callout height.</summary>
        public double? Height { get; set; }

        internal Action<VisioCalloutOptions>? CreateOptionsConfigurator() {
            if (!Width.HasValue && !Height.HasValue) {
                return null;
            }

            return options => {
                if (Width.HasValue) {
                    options.Width = Width.Value;
                }

                if (Height.HasValue) {
                    options.Height = Height.Value;
                }
            };
        }
    }
}
