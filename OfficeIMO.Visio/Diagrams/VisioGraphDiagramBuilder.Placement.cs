using System;
using System.Linq;

namespace OfficeIMO.Visio.Diagrams {
    public sealed partial class VisioGraphDiagramBuilder {
        private bool _preserveLayout;

        /// <summary>
        /// Uses imported node and container placements without automatic layout or geometry polish.
        /// Every node must supply a placement. Imported routes keep their explicit glued endpoints.
        /// </summary>
        public VisioGraphDiagramBuilder PreserveLayout(bool enabled = true) {
            _preserveLayout = enabled;
            return this;
        }

        private void AssignPreservedCoordinates() {
            foreach (NodeItem node in _nodes) {
                var placement = node.Placement ?? throw new InvalidOperationException("Node '" + node.Id + "' requires explicit bounds when preserving layout.");
                node.PinX = placement.PinX; node.PinY = placement.PinY;
                if (_fitPageToGraph) {
                    _pageWidth = Math.Max(_pageWidth, placement.PinX + placement.Width / 2 + _rightMargin);
                    _pageHeight = Math.Max(_pageHeight, placement.PinY + placement.Height / 2 + _topMargin);
                }
            }
            if (_fitPageToGraph) foreach (ZoneItem zone in _zones) {
                GetZoneBounds(zone, out _, out _, out double right, out double top);
                _pageWidth = Math.Max(_pageWidth, right + _rightMargin);
                _pageHeight = Math.Max(_pageHeight, top + _topMargin);
            }
        }

        private void ApplyPreservedRoute(VisioConnector connector, VisioGraphRoute route) {
            connector.Kind = ConnectorKind.Straight;
            connector.RerouteBehavior = VisioConnectorRerouteBehavior.Never;
            var first = route.Points[0];
            var last = route.Points[route.Points.Count - 1];
            connector.FromConnectionPoint = AddRouteEndpoint(connector.From, first.X, first.Y);
            connector.ToConnectionPoint = AddRouteEndpoint(connector.To, last.X, last.Y);
            connector.Waypoints.Clear();
            for (int i = 1; i < route.Points.Count - 1; i++) {
                var point = route.Points[i];
                connector.Waypoints.Add(new VisioConnectorWaypoint(point.X.ToInches(_unit), point.Y.ToInches(_unit)));
            }
        }

        private VisioConnectionPoint AddRouteEndpoint(VisioShape shape, double x, double y) {
            double localX = x.ToInches(_unit) - shape.PinX + shape.Width / 2;
            double localY = y.ToInches(_unit) - shape.PinY + shape.Height / 2;
            var existing = shape.ConnectionPoints.FirstOrDefault(point => Math.Abs(point.X - localX) < 1e-9 && Math.Abs(point.Y - localY) < 1e-9);
            if (existing != null) return existing;
            var created = new VisioConnectionPoint(localX, localY, 0, 0);
            shape.ConnectionPoints.Add(created);
            return created;
        }
    }
}
