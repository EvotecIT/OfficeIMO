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
        private void AddZones(VisioPage page) {
            foreach (ZoneItem zone in _zones) {
                GetZoneBounds(zone, out double left, out double bottom, out double right, out double top);
                double width = right - left;
                double height = top - bottom;
                VisioShape shape = VisioNetworkDiagramVisuals.CreateBackgroundZone(
                    _document,
                    zone.Id,
                    left + width / 2D,
                    bottom + height / 2D,
                    width,
                    height,
                    string.Empty,
                    _theme,
                    _unit);
                page.Shapes.Add(shape);
                ApplyZoneMetadata(shape, zone);
                VisioNetworkDiagramVisuals.AddBackgroundZoneCaption(
                    page,
                    CreateGeneratedId(VisioNetworkDiagramVisuals.CreateBackgroundZoneCaptionId(zone.Id)),
                    zone.Text,
                    left,
                    top,
                    width,
                    _theme);
            }
        }

        private void AddNodes(VisioPage page) {
            foreach (NodeItem node in _nodes) {
                GetNodeShape(node, out string masterNameU, out double width, out double height);
                VisioShape shape;
                if (node.Stencil != null) {
                    shape = page.AddStencilShape(node.Stencil, node.Id, node.PinX, node.PinY, width, height, string.Empty, node.StencilCatalogName);
                    VisioShapeStyle? stencilStyle = node.StyleOverride ?? GetBuiltInStencilNodeStyle(node);
                    stencilStyle?.ApplyTo(shape);
                } else {
                    shape = new VisioShape(node.Id, node.PinX.ToInches(_unit), node.PinY.ToInches(_unit), width.ToInches(_unit), height.ToInches(_unit), node.Text) {
                        NameU = masterNameU,
                    };
                    (node.StyleOverride ?? GetNodeStyle(node.Kind)).ApplyTo(shape);
                    page.Shapes.Add(shape);
                }

                node.Shape = shape;
                ApplyNodeMetadata(shape, node);
                if (node.Stencil != null && !string.IsNullOrWhiteSpace(node.Text)) {
                    AddStencilNodeCaption(page, node, width, height);
                }
            }
        }

        private static void ApplyNodeMetadata(VisioShape shape, NodeItem node) {
            foreach (NodeShapeDataItem data in node.ShapeData) {
                shape.SetShapeData(data.Name, data.Value, data.Label, data.Type, data.Prompt, data.Format);
            }

            foreach (VisioHyperlink hyperlink in node.Hyperlinks) {
                VisioHyperlink target = shape.AddHyperlink(hyperlink.Address ?? string.Empty, hyperlink.Description, hyperlink.SubAddress);
                CopyHyperlinkSettings(hyperlink, target);
            }
        }

        private static void CopyHyperlinkSettings(VisioHyperlink source, VisioHyperlink target) {
            target.RowName = source.RowName;
            target.ExtraInfo = source.ExtraInfo;
            target.Frame = source.Frame;
            target.NewWindow = source.NewWindow;
            target.Default = source.Default;
            target.Invisible = source.Invisible;
            target.SortKey = source.SortKey;
        }

        private void AddStencilNodeCaption(VisioPage page, NodeItem node, double width, double height) {
            VisioShape label = page.AddTextBox(
                CreateGeneratedId(node.Id + "-label"),
                node.PinX,
                node.PinY - (height / 2D) - 0.26D,
                Math.Max(1.15D, width + 0.55D),
                0.34D,
                node.Text);
            label.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 9.5D,
                Color = Color.FromRgb(25, 35, 45),
                BackgroundColor = Color.White,
                BackgroundTransparency = 0,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
            MarkDiagramAdornment(label);
        }

        private void AddEdges(VisioPage page) {
            int routeIndex = 0;
            HashSet<string> reservedConnectorIds = BuildReservedConnectorIds(page);
            foreach (EdgeItem edge in _edges) {
                NodeItem from = _nodesById[edge.FromId];
                NodeItem to = _nodesById[edge.ToId];
                if (from.Shape == null || to.Shape == null) {
                    throw new InvalidOperationException("Nodes must be placed before graph edges are created.");
                }

                VisioNetworkDiagramVisuals.ResolveSides(from.Shape, to.Shape, out VisioSide fromSide, out VisioSide toSide);
                ConnectorKind connectorKind = _layout == VisioGraphLayout.Radial ? ConnectorKind.Straight : ConnectorKind.RightAngle;
                string connectorId = edge.Id ?? ReserveGeneratedConnectorId(reservedConnectorIds);
                VisioConnector connector = page.AddConnector(connectorId, from.Shape, to.Shape, connectorKind, fromSide, toSide);

                (edge.StyleOverride ?? GetConnectorStyle(edge.Kind, edge.Directed)).ApplyTo(connector);
                connector.Label = edge.Label;
                ApplyEdgeMetadata(connector, edge);
                if (_layout != VisioGraphLayout.Radial) {
                    connector.RouteOrthogonal(offset: (routeIndex % 7) * 0.05D);
                }

                if (!string.IsNullOrWhiteSpace(edge.Label)) {
                    VisioTextStyle labelStyle = connector.TextStyle?.Clone() ?? new VisioTextStyle();
                    labelStyle.BackgroundColor = Color.White;
                    labelStyle.BackgroundTransparency = 0;
                    labelStyle.Size = Math.Max(labelStyle.Size ?? 0D, 8.5D);
                    connector.TextStyle = labelStyle;
                    connector.PlaceLabel(0.5D, offsetY: 0.26D, width: 1.35D, height: 0.32D);
                    connector.ResizeLabelToText(maximumWidth: 1.5D);
                }

                routeIndex++;
            }
        }

        private HashSet<string> BuildReservedConnectorIds(VisioPage page) {
            HashSet<string> ids = new(StringComparer.OrdinalIgnoreCase);
            foreach (VisioShape shape in page.Shapes) {
                ReserveShapeIds(shape, ids);
            }

            foreach (VisioConnector connector in page.Connectors) {
                ids.Add(connector.Id);
            }

            foreach (EdgeItem edge in _edges) {
                if (edge.Id != null) {
                    ids.Add(edge.Id);
                }
            }

            return ids;
        }

        private static void ReserveShapeIds(VisioShape shape, HashSet<string> ids) {
            ids.Add(shape.Id);
            foreach (VisioShape child in shape.Children) {
                ReserveShapeIds(child, ids);
            }
        }

        private static string ReserveGeneratedConnectorId(HashSet<string> reservedIds) {
            int nextId = 1;
            while (true) {
                string candidate = nextId.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (reservedIds.Add(candidate)) {
                    return candidate;
                }

                nextId++;
            }
        }

        private static void ApplyEdgeMetadata(VisioConnector connector, EdgeItem edge) {
            foreach (NodeShapeDataItem data in edge.ShapeData) {
                connector.SetShapeData(data.Name, data.Value, data.Label, data.Type, data.Prompt, data.Format);
            }

            foreach (VisioHyperlink hyperlink in edge.Hyperlinks) {
                VisioHyperlink target = connector.AddHyperlink(hyperlink.Address ?? string.Empty, hyperlink.Description, hyperlink.SubAddress);
                CopyHyperlinkSettings(hyperlink, target);
            }
        }
    }
}
