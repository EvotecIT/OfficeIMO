using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Layout and geometry helpers for Visio pages, shapes, and selections.
    /// </summary>
    public static partial class VisioLayoutExtensions {
        private readonly struct Point {
            public Point(double x, double y) {
                X = x;
                Y = y;
            }

            public double X { get; }

            public double Y { get; }

            public Point Offset(double x, double y) {
                return new Point(X + x, Y + y);
            }
        }

        private readonly struct LabelCandidate {
            public LabelCandidate(double offsetX, double offsetY, double positionDelta) {
                OffsetX = offsetX;
                OffsetY = offsetY;
                PositionDelta = positionDelta;
            }

            public double OffsetX { get; }

            public double OffsetY { get; }

            public double PositionDelta { get; }
        }

        private readonly struct ShapeCandidate {
            public ShapeCandidate(double offsetX, double offsetY) {
                OffsetX = offsetX;
                OffsetY = offsetY;
            }

            public double OffsetX { get; }

            public double OffsetY { get; }
        }

        private readonly struct ConnectorLabelBounds {
            public ConnectorLabelBounds(VisioConnector connector, VisioShapeBounds bounds) {
                Connector = connector;
                Bounds = bounds;
            }

            public VisioConnector Connector { get; }

            public VisioShapeBounds Bounds { get; }
        }

        private readonly struct ConnectorLabelWorkItem {
            public ConnectorLabelWorkItem(VisioConnector connector, CandidateScore score, int index) {
                Connector = connector;
                Score = score;
                Index = index;
            }

            public VisioConnector Connector { get; }

            public CandidateScore Score { get; }

            public int Index { get; }
        }

        private readonly struct CandidateScore {
            public CandidateScore(double pagePenalty, double shapeOverlap, double labelOverlap, double connectorPathOverlap, double zonePenalty) {
                PagePenalty = pagePenalty;
                ShapeOverlap = shapeOverlap;
                LabelOverlap = labelOverlap;
                ConnectorPathOverlap = connectorPathOverlap;
                ZonePenalty = zonePenalty;
            }

            public double PagePenalty { get; }

            public double ShapeOverlap { get; }

            public double LabelOverlap { get; }

            public double ConnectorPathOverlap { get; }

            public double ZonePenalty { get; }

            public double TotalPenalty => PagePenalty + ShapeOverlap + LabelOverlap + ConnectorPathOverlap + ZonePenalty;

            public bool HasConflict => PagePenalty > 1e-9 || ShapeOverlap > 1e-9 || LabelOverlap > 1e-9 || ConnectorPathOverlap > 1e-9;

            public bool HasImprovementOpportunity => HasConflict || ZonePenalty > 1e-9;

            public bool IsBetterThan(CandidateScore other) {
                if (PagePenalty < other.PagePenalty - 1e-9) {
                    return true;
                }

                if (PagePenalty > other.PagePenalty + 1e-9) {
                    return false;
                }

                if (ShapeOverlap < other.ShapeOverlap - 1e-9) {
                    return true;
                }

                if (ShapeOverlap > other.ShapeOverlap + 1e-9) {
                    return false;
                }

                if (LabelOverlap < other.LabelOverlap - 1e-9) {
                    return true;
                }

                if (LabelOverlap > other.LabelOverlap + 1e-9) {
                    return false;
                }

                if (ConnectorPathOverlap < other.ConnectorPathOverlap - 1e-9) {
                    return true;
                }

                if (ConnectorPathOverlap > other.ConnectorPathOverlap + 1e-9) {
                    return false;
                }

                return ZonePenalty < other.ZonePenalty - 1e-9;
            }
        }
    }
}
