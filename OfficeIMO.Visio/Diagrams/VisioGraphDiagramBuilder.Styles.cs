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
        private void GetZoneBounds(ZoneItem zone, out double left, out double bottom, out double right, out double top) {
            const double horizontalPadding = 0.45D;
            const double verticalPadding = 0.35D;
            left = double.MaxValue;
            bottom = double.MaxValue;
            right = double.MinValue;
            top = double.MinValue;
            foreach (string nodeId in zone.NodeIds) {
                NodeItem node = _nodesById[nodeId];
                GetNodeShape(node, out _, out double width, out double height);
                left = Math.Min(left, node.PinX - width / 2D);
                bottom = Math.Min(bottom, node.PinY - height / 2D);
                right = Math.Max(right, node.PinX + width / 2D);
                top = Math.Max(top, node.PinY + height / 2D);
                if (HasStencilCaption(node)) {
                    bottom = Math.Min(bottom, node.PinY - height / 2D - StencilCaptionBottomOverflow);
                }
            }

            left -= horizontalPadding;
            bottom -= verticalPadding;
            right += horizontalPadding;
            top += verticalPadding;
        }

        private double XForLayer(int layer) {
            double width = LayoutNodeWidth();
            return _leftMargin + (width / 2D) + layer * (width + _columnGap);
        }

        private double XForRow(int row) {
            double width = LayoutNodeWidth();
            double contentWidth = _maximumRows * width + Math.Max(0, _maximumRows - 1) * _columnGap;
            double availableWidth = _pageWidth - _leftMargin - _rightMargin;
            double start = _leftMargin + Math.Max(0D, (availableWidth - contentWidth) / 2D);
            return start + (width / 2D) + row * (width + _columnGap);
        }

        private double YForLayer(int layer) {
            double height = LayoutNodeHeight();
            double top = _pageHeight - _topMargin - HeaderHeight;
            return top - (height / 2D) - layer * (height + _rowGap);
        }

        private double YForRow(int row) {
            double height = LayoutNodeHeight();
            double contentHeight = _maximumRows * height + Math.Max(0, _maximumRows - 1) * _rowGap;
            double top = _pageHeight - _topMargin - HeaderHeight;
            double availableHeight = _pageHeight - _topMargin - _bottomMargin - HeaderHeight;
            double layerTop = top - Math.Max(0D, (availableHeight - contentHeight) / 2D);
            return layerTop - (height / 2D) - row * (height + _rowGap);
        }

        private double LayoutNodeWidth() {
            double width = _nodeWidth;
            foreach (NodeItem node in _nodes) {
                GetNodeShape(node, out _, out double nodeWidth, out _);
                width = Math.Max(width, nodeWidth);
            }

            return width;
        }

        private double LayoutNodeHeight() {
            double height = _nodeHeight;
            foreach (NodeItem node in _nodes) {
                GetNodeShape(node, out _, out _, out double nodeHeight);
                height = Math.Max(height, nodeHeight);
            }

            return height;
        }

        private double HeaderHeight {
            get {
                double height = TitleHeaderHeight + LegendHeaderHeight;
                if (_zones.Any(zone => !string.IsNullOrWhiteSpace(zone.Text))) {
                    height += VisioNetworkDiagramVisuals.BackgroundZoneCaptionHeaderClearance;
                }

                return height;
            }
        }

        private static void ApplyZoneMetadata(VisioShape shape, ZoneItem zone) {
            foreach (NodeShapeDataItem data in zone.ShapeData) {
                shape.SetShapeData(data.Name, data.Value, data.Label, data.Type, data.Prompt, data.Format);
            }

            foreach (VisioHyperlink hyperlink in zone.Hyperlinks) {
                VisioHyperlink target = shape.AddHyperlink(hyperlink.Address ?? string.Empty, hyperlink.Description, hyperlink.SubAddress);
                CopyHyperlinkSettings(hyperlink, target);
            }
        }

        private VisioShapeStyle? GetBuiltInStencilNodeStyle(NodeItem node) {
            if (node.Stencil == null || !string.IsNullOrWhiteSpace(node.Stencil.SourcePackagePath)) {
                return null;
            }

            return GetNodeStyle(node.Kind);
        }

        private static bool HasStencilCaption(NodeItem node) {
            return node.Stencil != null && !string.IsNullOrWhiteSpace(node.Text);
        }

        private void GetNodeShape(NodeItem node, out string masterNameU, out double width, out double height) {
            width = node.Stencil?.DefaultWidth ?? _nodeWidth;
            height = node.Stencil?.DefaultHeight ?? _nodeHeight;
            if (node.Stencil != null) {
                masterNameU = node.Stencil.MasterNameU;
                VisioMeasurementUnit sourceUnit = node.Stencil.DefaultUnit ?? _unit;
                width = width.ToInches(sourceUnit).FromInches(_unit);
                height = height.ToInches(sourceUnit).FromInches(_unit);
                return;
            }

            switch (node.Kind) {
                case VisioGraphNodeKind.Data:
                    masterNameU = "Data";
                    break;
                case VisioGraphNodeKind.External:
                    masterNameU = "Circle";
                    width = Math.Min(_nodeWidth, _nodeHeight * 1.15D);
                    height = width;
                    break;
                case VisioGraphNodeKind.Decision:
                    masterNameU = "Decision";
                    height = _nodeHeight * 1.2D;
                    break;
                default:
                    masterNameU = "Process";
                    break;
            }
        }

        private VisioShapeStyle GetNodeStyle(VisioGraphNodeKind kind) {
            switch (kind) {
                case VisioGraphNodeKind.Data:
                    return _theme.Marker;
                case VisioGraphNodeKind.External:
                    return _theme.Success;
                case VisioGraphNodeKind.Decision:
                    return _theme.Decision;
                case VisioGraphNodeKind.Emphasis:
                    return _theme.Emphasis;
                default:
                    return _theme.Primary;
            }
        }

        private static VisioGraphNodeKind InferStencilNodeKind(VisioStencilShape stencil) {
            string masterName = stencil.MasterNameU ?? string.Empty;
            if (string.Equals(masterName, "Data", StringComparison.OrdinalIgnoreCase)) {
                return VisioGraphNodeKind.Data;
            }

            if (string.Equals(masterName, "Decision", StringComparison.OrdinalIgnoreCase)) {
                return VisioGraphNodeKind.Decision;
            }

            if (string.Equals(masterName, "Circle", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(stencil.Category, "Network", StringComparison.OrdinalIgnoreCase) && string.Equals(stencil.Id, "net.internet", StringComparison.OrdinalIgnoreCase)) {
                return VisioGraphNodeKind.External;
            }

            return VisioGraphNodeKind.Process;
        }

        private VisioConnectorStyle GetConnectorStyle(VisioGraphConnectorKind kind, bool directed) {
            VisioConnectorStyle style;
            switch (kind) {
                case VisioGraphConnectorKind.Data:
                    style = _theme.DataConnector.Clone();
                    break;
                case VisioGraphConnectorKind.Control:
                    style = _theme.ControlConnector.Clone();
                    break;
                case VisioGraphConnectorKind.Emphasis:
                    style = _theme.Connector.Clone();
                    style.LineWeight = Math.Max(style.LineWeight, 0.026D);
                    break;
                default:
                    style = _theme.Connector.Clone();
                    break;
            }

            if (!directed) {
                style.EndArrow = EndArrow.None;
            }

            return style;
        }

        private static string GetNodeKindLabel(VisioGraphNodeKind kind) {
            switch (kind) {
                case VisioGraphNodeKind.Data:
                    return "Data store";
                case VisioGraphNodeKind.Decision:
                    return "Decision";
                case VisioGraphNodeKind.Emphasis:
                    return "Emphasis";
                case VisioGraphNodeKind.External:
                    return "External";
                default:
                    return "Process";
            }
        }

        private static string GetConnectorLegendLabel(VisioGraphConnectorKind kind, bool directed) {
            if (!directed) {
                switch (kind) {
                    case VisioGraphConnectorKind.Data:
                        return "Data relationship";
                    case VisioGraphConnectorKind.Control:
                        return "Control relationship";
                    case VisioGraphConnectorKind.Emphasis:
                        return "Emphasized relationship";
                    default:
                        return "Relationship";
                }
            }

            switch (kind) {
                case VisioGraphConnectorKind.Data:
                    return "Data flow";
                case VisioGraphConnectorKind.Control:
                    return "Control flow";
                case VisioGraphConnectorKind.Emphasis:
                    return "Emphasized flow";
                default:
                    return "Dependency";
            }
        }
    }
}
