using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    public static partial class VisioConnectorRoutingExtensions {
        private readonly struct RoutePoint {
            public RoutePoint(double x, double y) {
                X = x;
                Y = y;
            }

            public double X { get; }

            public double Y { get; }
        }

        private readonly struct RouteCandidate {
            public RouteCandidate(params RoutePoint[] points)
                : this(points, new RouteScore(int.MaxValue, int.MaxValue, double.PositiveInfinity)) {
            }

            private RouteCandidate(IReadOnlyList<RoutePoint> points, RouteScore score) {
                if (points.Count < 2) {
                    throw new ArgumentException("Route candidates require at least two points.", nameof(points));
                }

                Points = CollapseDuplicatePoints(points);
                Score = score;
            }

            public IReadOnlyList<RoutePoint> Points { get; }

            public IReadOnlyList<RoutePoint> Waypoints => Points.Skip(1).Take(Points.Count - 2).ToList();

            public RouteScore Score { get; }

            public double Length {
                get {
                    double length = 0D;
                    for (int i = 1; i < Points.Count; i++) {
                        length += Distance(Points[i - 1], Points[i]);
                    }

                    return length;
                }
            }

            public RouteCandidate WithScore(RouteScore score) {
                return new RouteCandidate(Points, score);
            }

            private static IReadOnlyList<RoutePoint> CollapseDuplicatePoints(IReadOnlyList<RoutePoint> points) {
                List<RoutePoint> collapsed = new();
                foreach (RoutePoint point in points) {
                    if (collapsed.Count == 0 || !PointsEqual(collapsed[collapsed.Count - 1], point)) {
                        collapsed.Add(point);
                    }
                }

                return collapsed;
            }
        }

        private readonly struct ConnectorRoutingWorkItem {
            public ConnectorRoutingWorkItem(VisioConnector connector, RouteScore score, int index) {
                Connector = connector;
                Score = score;
                Index = index;
            }

            public VisioConnector Connector { get; }

            public RouteScore Score { get; }

            public int Index { get; }
        }

        private readonly struct RouteScore {
            public RouteScore(int intersections, int connectorCrossings, double length) {
                Intersections = intersections;
                ConnectorCrossings = connectorCrossings;
                Length = length;
            }

            public int Intersections { get; }

            public int ConnectorCrossings { get; }

            public double Length { get; }

            public bool HasConflicts => Intersections > 0 || ConnectorCrossings > 0;

            public bool IsBetterThan(RouteScore other) {
                if (Intersections != other.Intersections) {
                    return Intersections < other.Intersections;
                }

                if (ConnectorCrossings != other.ConnectorCrossings) {
                    return ConnectorCrossings < other.ConnectorCrossings;
                }

                return Length < other.Length - 1e-9;
            }
        }

        private static double Distance(RoutePoint from, RoutePoint to) {
            return OfficeGeometry.Distance((from.X, from.Y), (to.X, to.Y));
        }
    }
}
