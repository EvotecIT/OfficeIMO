using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio.Diagrams {
    public sealed partial class VisioGraphDiagramBuilder {
        private bool _preserveLayout;
        private readonly Dictionary<VisioShape, RouteEndpointIndex> _routeEndpointIndex = new();
        private const double RouteEndpointTolerance = 1e-9;

        private sealed class RouteEndpointIndex {
            internal readonly Dictionary<(long X, long Y), List<VisioConnectionPoint>> Buckets = new();
            internal int IndexedCount;
        }

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
            if (!_routeEndpointIndex.TryGetValue(shape, out RouteEndpointIndex? index)) {
                index = new RouteEndpointIndex();
                _routeEndpointIndex.Add(shape, index);
            }
            // AddConnector can append side points after this shape was indexed for an earlier edge.
            for (int i = index.IndexedCount; i < shape.ConnectionPoints.Count; i++)
                AddIndexedPoint(index.Buckets, shape.ConnectionPoints[i]);
            index.IndexedCount = shape.ConnectionPoints.Count;
            var buckets = index.Buckets;
            (long X, long Y) bucket = RouteEndpointBucket(localX, localY);
            for (long dx = -1; dx <= 1; dx++) for (long dy = -1; dy <= 1; dy++) {
                if (!buckets.TryGetValue((bucket.X + dx, bucket.Y + dy), out var points)) continue;
                foreach (VisioConnectionPoint point in points) {
                    if (Math.Abs(point.X - localX) < RouteEndpointTolerance &&
                        Math.Abs(point.Y - localY) < RouteEndpointTolerance) return point;
                }
            }
            var created = new VisioConnectionPoint(localX, localY, 0, 0);
            shape.ConnectionPoints.Add(created);
            AddIndexedPoint(buckets, created);
            index.IndexedCount++;
            return created;
        }

        private static void AddIndexedPoint(Dictionary<(long X, long Y), List<VisioConnectionPoint>> buckets,
            VisioConnectionPoint point) {
            var bucket = RouteEndpointBucket(point.X, point.Y);
            if (!buckets.TryGetValue(bucket, out var points)) buckets.Add(bucket, points = new List<VisioConnectionPoint>());
            points.Add(point);
        }

        private static (long X, long Y) RouteEndpointBucket(double x, double y) {
            double bucketX = Math.Floor(x / RouteEndpointTolerance);
            double bucketY = Math.Floor(y / RouteEndpointTolerance);
            if (double.IsNaN(bucketX) || double.IsNaN(bucketY) ||
                bucketX < long.MinValue + 1D || bucketX > long.MaxValue - 1D ||
                bucketY < long.MinValue + 1D || bucketY > long.MaxValue - 1D)
                throw new InvalidOperationException("A preserved route endpoint is outside the supported coordinate range.");
            return ((long)bucketX, (long)bucketY);
        }
    }
}
