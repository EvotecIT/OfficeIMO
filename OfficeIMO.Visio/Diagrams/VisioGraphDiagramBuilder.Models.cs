using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for generic graph diagrams where OfficeIMO lays out arbitrary nodes and edges.
    /// </summary>
    public sealed partial class VisioGraphDiagramBuilder {
        private sealed class NodeItem {
            public NodeItem(string id, string text, VisioGraphNodeKind kind, VisioStencilShape? stencil, string? stencilCatalogName) {
                Id = id;
                Text = text;
                Kind = kind;
                Stencil = stencil;
                StencilCatalogName = stencilCatalogName;
            }

            public string Id { get; }

            public string Text { get; }

            public VisioGraphNodeKind Kind { get; }

            public VisioStencilShape? Stencil { get; }

            public string? StencilCatalogName { get; }

            public int Layer { get; set; }

            public int Row { get; set; }

            public double PinX { get; set; }

            public double PinY { get; set; }

            public VisioShape? Shape { get; set; }

            public VisioShapeStyle? StyleOverride { get; set; }

            public List<NodeShapeDataItem> ShapeData { get; } = new();

            public List<VisioHyperlink> Hyperlinks { get; } = new();
        }

        private sealed class NodeShapeDataItem {
            public NodeShapeDataItem(string name, string? value, string? label, VisioShapeDataType? type, string? prompt, string? format) {
                Name = name;
                Value = value;
                Label = label;
                Type = type;
                Prompt = prompt;
                Format = format;
            }

            public string Name { get; }

            public string? Value { get; }

            public string? Label { get; }

            public VisioShapeDataType? Type { get; }

            public string? Prompt { get; }

            public string? Format { get; }
        }

        private sealed class EdgeItem {
            public EdgeItem(string? id, string fromId, string toId, VisioGraphConnectorKind kind, string? label, bool directed) {
                Id = id;
                FromId = fromId;
                ToId = toId;
                Kind = kind;
                Label = label;
                Directed = directed;
            }

            public string? Id { get; }

            public string FromId { get; }

            public string ToId { get; }

            public VisioGraphConnectorKind Kind { get; }

            public string? Label { get; }

            public bool Directed { get; }

            public VisioConnectorStyle? StyleOverride { get; set; }

            public List<NodeShapeDataItem> ShapeData { get; } = new();

            public List<VisioHyperlink> Hyperlinks { get; } = new();
        }

        private sealed class ZoneItem {
            public ZoneItem(string id, string text, IReadOnlyList<string> nodeIds) {
                Id = id;
                Text = text;
                NodeIds = nodeIds;
            }

            public string Id { get; }

            public string Text { get; }

            public IReadOnlyList<string> NodeIds { get; }

            public List<NodeShapeDataItem> ShapeData { get; } = new();

            public List<VisioHyperlink> Hyperlinks { get; } = new();
        }

        private sealed class LegendItem {
            public LegendItem(string idSuffix, string label, VisioGraphNodeKind? nodeKind, VisioGraphConnectorKind? connectorKind, bool directed) {
                IdSuffix = idSuffix;
                Label = label;
                NodeKind = nodeKind;
                ConnectorKind = connectorKind;
                Directed = directed;
            }

            public string IdSuffix { get; }

            public string Label { get; }

            public VisioGraphNodeKind? NodeKind { get; }

            public VisioGraphConnectorKind? ConnectorKind { get; }

            public bool Directed { get; }
        }
    }
}
