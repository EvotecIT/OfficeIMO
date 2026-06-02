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
        internal VisioPage Build() {
            if (_built) {
                throw new InvalidOperationException("This graph diagram builder has already produced a page.");
            }

            _built = true;
            if (_nodes.Count == 0) {
                throw new InvalidOperationException("A graph diagram requires at least one node.");
            }

            ValidateZones();
            AssignLayoutMetadata();
            SizePageForLayout();
            AssignCoordinates();

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddZones(page);
            AddNodes(page);
            AddEdges(page);
            AddLegend(page);
            AddTitle(page);
            page.PolishDiagram(new VisioDiagramPolishOptions {
                FitToContent = false,
                ResizeShapesToText = false,
                ResizeConnectorLabelsToText = true,
                ResolveConnectorShapeIntersections = _layout != VisioGraphLayout.Radial,
                ResolveConnectorLabelOverlaps = true
            });
            _document.RequestRecalcOnOpen();
            return page;
        }

        private void ValidateZones() {
            foreach (ZoneItem zone in _zones) {
                foreach (string nodeId in zone.NodeIds) {
                    EnsureKnownNode(nodeId, nameof(zone.NodeIds));
                }
            }
        }

        private void AssignLayoutMetadata() {
            if (_layout == VisioGraphLayout.Grid) {
                AssignGridMetadata();
                return;
            }

            AssignBreadthFirstMetadata();
        }

        private void AssignGridMetadata() {
            int columns = Math.Max(1, (int)Math.Ceiling(Math.Sqrt(_nodes.Count)));
            for (int i = 0; i < _nodes.Count; i++) {
                _nodes[i].Layer = i % columns;
                _nodes[i].Row = i / columns;
            }

            _maximumRows = _nodes.GroupBy(node => node.Layer).Max(group => group.Count());
        }

        private void AssignBreadthFirstMetadata() {
            Dictionary<string, List<string>> outgoing = _nodes.ToDictionary(node => node.Id, _ => new List<string>(), StringComparer.Ordinal);
            Dictionary<string, List<string>> undirected = _nodes.ToDictionary(node => node.Id, _ => new List<string>(), StringComparer.Ordinal);
            Dictionary<string, int> indegree = _nodes.ToDictionary(node => node.Id, _ => 0, StringComparer.Ordinal);
            foreach (EdgeItem edge in _edges) {
                outgoing[edge.FromId].Add(edge.ToId);
                undirected[edge.FromId].Add(edge.ToId);
                undirected[edge.ToId].Add(edge.FromId);
                if (edge.Directed) {
                    indegree[edge.ToId]++;
                } else {
                    outgoing[edge.ToId].Add(edge.FromId);
                }
            }

            HashSet<string> assigned = new(StringComparer.Ordinal);
            Queue<NodeItem> ready = new();

            void Enqueue(NodeItem node, int layer) {
                if (assigned.Add(node.Id)) {
                    node.Layer = layer;
                    ready.Enqueue(node);
                }
            }

            if (_rootIds.Count > 0) {
                foreach (string rootId in _rootIds) {
                    Enqueue(_nodesById[rootId], 0);
                }
            } else {
                foreach (NodeItem root in _nodes.Where(node => indegree[node.Id] == 0)) {
                    Enqueue(root, 0);
                }
            }

            if (assigned.Count == 0 && _nodes.Count > 0) {
                Enqueue(_nodes[0], 0);
            }

            while (assigned.Count < _nodes.Count || ready.Count > 0) {
                while (ready.Count > 0) {
                    NodeItem node = ready.Dequeue();
                    IReadOnlyList<string> nextIds = outgoing[node.Id].Count > 0 ? outgoing[node.Id] : undirected[node.Id];
                    foreach (string nextId in nextIds) {
                        if (!assigned.Contains(nextId)) {
                            Enqueue(_nodesById[nextId], node.Layer + 1);
                        }
                    }
                }

                if (assigned.Count < _nodes.Count) {
                    NodeItem nextRoot = _nodes.First(node => !assigned.Contains(node.Id));
                    Enqueue(nextRoot, 0);
                }
            }

            foreach (IGrouping<int, NodeItem> layer in _nodes.GroupBy(node => node.Layer).OrderBy(group => group.Key)) {
                int row = 0;
                foreach (NodeItem node in layer.OrderBy(node => _nodes.IndexOf(node))) {
                    node.Row = row;
                    row++;
                }
            }

            _maximumRows = Math.Max(1, _nodes.GroupBy(node => node.Layer).Max(group => group.Count()));
        }

        private void SizePageForLayout() {
            if (!_fitPageToGraph) {
                return;
            }

            int layerCount = Math.Max(1, _nodes.Max(node => node.Layer) + 1);
            int rowCount = Math.Max(1, _nodes.GroupBy(node => node.Layer).Max(group => group.Count()));
            double layoutNodeWidth = LayoutNodeWidth();
            double layoutNodeHeight = LayoutNodeHeight();
            double requiredWidth;
            double requiredHeight;
            if (_layout == VisioGraphLayout.Radial) {
                double radius = Math.Max(1D, _nodes.Max(node => node.Layer)) * Math.Max(layoutNodeWidth + _columnGap, layoutNodeHeight + _rowGap);
                requiredWidth = _leftMargin + _rightMargin + (radius * 2D) + layoutNodeWidth * 2D;
                requiredHeight = _topMargin + _bottomMargin + HeaderHeight + (radius * 2D) + layoutNodeHeight * 2D;
            } else if (_direction == VisioGraphDirection.TopToBottom) {
                requiredWidth = _leftMargin + _rightMargin + (rowCount * layoutNodeWidth) + Math.Max(0, rowCount - 1) * _columnGap;
                requiredHeight = _topMargin + _bottomMargin + HeaderHeight + (layerCount * layoutNodeHeight) + Math.Max(0, layerCount - 1) * _rowGap;
            } else {
                requiredWidth = _leftMargin + _rightMargin + (layerCount * layoutNodeWidth) + Math.Max(0, layerCount - 1) * _columnGap;
                requiredHeight = _topMargin + _bottomMargin + HeaderHeight + (rowCount * layoutNodeHeight) + Math.Max(0, rowCount - 1) * _rowGap;
            }

            if (_nodes.Any(HasStencilCaption)) {
                requiredHeight += StencilCaptionBottomOverflow;
            }

            _pageWidth = Math.Max(_pageWidth, requiredWidth);
            _pageHeight = Math.Max(_pageHeight, requiredHeight);
        }

        private void AssignCoordinates() {
            if (_layout == VisioGraphLayout.Radial) {
                AssignRadialCoordinates();
                return;
            }

            foreach (NodeItem node in _nodes) {
                if (_direction == VisioGraphDirection.TopToBottom) {
                    node.PinX = XForRow(node.Row);
                    node.PinY = YForLayer(node.Layer);
                } else {
                    node.PinX = XForLayer(node.Layer);
                    node.PinY = YForRow(node.Row);
                }
            }
        }

        private void AssignRadialCoordinates() {
            double contentHeight = _pageHeight - _topMargin - _bottomMargin - HeaderHeight;
            double centerX = _leftMargin + ((_pageWidth - _leftMargin - _rightMargin) / 2D);
            double centerY = _bottomMargin + (contentHeight / 2D);
            double ringGap = Math.Max(LayoutNodeWidth() + _columnGap, LayoutNodeHeight() + _rowGap);
            foreach (IGrouping<int, NodeItem> layer in _nodes.GroupBy(node => node.Layer).OrderBy(group => group.Key)) {
                NodeItem[] layerNodes = layer.OrderBy(node => _nodes.IndexOf(node)).ToArray();
                double radius = layer.Key == 0 && layerNodes.Length == 1 ? 0D : Math.Max(0.9D, layer.Key) * ringGap;
                for (int i = 0; i < layerNodes.Length; i++) {
                    double angle = layerNodes.Length == 1 ? -Math.PI / 2D : (-Math.PI / 2D) + (2D * Math.PI * i / layerNodes.Length);
                    layerNodes[i].PinX = centerX + Math.Cos(angle) * radius;
                    layerNodes[i].PinY = centerY + Math.Sin(angle) * radius;
                }
            }
        }
    }
}
