using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for UML-style sequence diagrams with deterministic participant,
    /// lifeline, and message placement.
    /// </summary>
    public sealed partial class VisioSequenceDiagramBuilder {
        private void AddNotes(VisioPage page, double firstMessageY) {
            List<LayoutBounds> reservedBounds = BuildNoteReservedBounds(page);
            foreach (NoteItem note in _notes) {
                NotePlacement placement = ResolveNotePlacement(page, note, firstMessageY, reservedBounds);

                VisioShape shape = page.AddStencilShape(VisioStencils.Sequence, "seq.note", note.Id, placement.X, placement.Y, NoteWidth, NoteHeight, note.Text);
                _theme.Container.ApplyTo(shape);
                shape.SetUserCell(VisioSemanticUserCells.Kind, "SequenceNote", "STR", prompt: "OfficeIMO semantic kind");
                shape.SetUserCell("OfficeIMO.SequenceParticipantId", note.ParticipantId, "STR", prompt: "OfficeIMO sequence note target participant");
                shape.SetUserCell("OfficeIMO.SequenceRowIndex", note.RowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence note row");
                shape.SetUserCell("OfficeIMO.SequenceRequestedPlacement", note.Placement.ToString(), "STR", prompt: "OfficeIMO requested sequence note placement");
                shape.SetUserCell("OfficeIMO.SequenceResolvedPlacement", placement.ResolvedPlacement.ToString(), "STR", prompt: "OfficeIMO resolved sequence note placement");
                page.AddToLayer(NoteLayer, shape);
                reservedBounds.Add(placement.Bounds.Inflate(NoteCollisionPadding));
            }
        }

        private void EnsureNotesFitPage(VisioPage page) {
            if (_notes.Count == 0) {
                return;
            }

            page.FitToContent(
                Math.Min(_leftMargin, _rightMargin).ToInches(_unit),
                Math.Min(_topMargin, _bottomMargin).ToInches(_unit));
        }

        private NotePlacement ResolveNotePlacement(VisioPage page, NoteItem note, double firstMessageY, IReadOnlyList<LayoutBounds> reservedBounds) {
            ParticipantItem participant = _participantsById[note.ParticipantId];
            double baseY = firstMessageY - (note.RowIndex * _messageSpacing) - (_messageSpacing / 2D);
            double minY = _bottomMargin + (NoteHeight / 2D);
            double pageHeight = page.Height.FromInches(_unit);
            double maxY = pageHeight - _topMargin - _participantHeight - NoteGap - (NoteHeight / 2D);
            VisioSide[] sides = note.Placement == VisioSide.Right
                ? new[] { VisioSide.Right, VisioSide.Left }
                : new[] { VisioSide.Left, VisioSide.Right };

            NotePlacement bestPlacement = default;
            double bestScore = double.PositiveInfinity;
            bool hasBest = false;
            foreach (VisioSide side in sides) {
                double x = GetNoteX(page, participant, side);
                for (int offsetIndex = 0; offsetIndex <= 12; offsetIndex++) {
                    foreach (double offset in GetNoteCandidateOffsets(offsetIndex)) {
                        double y = Math.Max(minY, Math.Min(maxY, baseY + offset));
                        LayoutBounds bounds = LayoutBounds.FromCenter(x, y, NoteWidth, NoteHeight);
                        double score = ScoreNotePlacement(bounds.Inflate(NoteCollisionPadding), reservedBounds) +
                                       (Math.Abs(offset) * 0.9D) +
                                       (side == note.Placement ? 0D : 1.8D);

                        if (score < bestScore) {
                            bestScore = score;
                            bestPlacement = new NotePlacement(x, y, side, bounds);
                            hasBest = true;
                        }

                        if (score <= 0.0001D) {
                            return bestPlacement;
                        }
                    }
                }
            }

            return hasBest
                ? bestPlacement
                : new NotePlacement(GetNoteX(page, participant, note.Placement), Math.Max(minY, Math.Min(maxY, baseY)), note.Placement, LayoutBounds.FromCenter(GetNoteX(page, participant, note.Placement), Math.Max(minY, Math.Min(maxY, baseY)), NoteWidth, NoteHeight));
        }

        private static IEnumerable<double> GetNoteCandidateOffsets(int offsetIndex) {
            if (offsetIndex == 0) {
                yield return 0D;
            } else {
                double offset = offsetIndex * NoteVerticalCandidateStep;
                yield return offset;
                yield return -offset;
            }
        }

        private double GetNoteX(VisioPage page, ParticipantItem participant, VisioSide placement) {
            double direction = placement == VisioSide.Right ? 1D : -1D;
            double x = participant.PinX + (direction * ((_participantWidth / 2D) + NoteGap + (NoteWidth / 2D)));
            double pageWidth = page.Width.FromInches(_unit);
            return Math.Max(_leftMargin + (NoteWidth / 2D), Math.Min(pageWidth - _rightMargin - (NoteWidth / 2D), x));
        }

        private static double ScoreNotePlacement(LayoutBounds candidate, IReadOnlyList<LayoutBounds> reservedBounds) {
            double score = 0D;
            foreach (LayoutBounds reserved in reservedBounds) {
                score += GetOverlapArea(candidate, reserved) * 25D;
            }

            return score;
        }

        private static double GetOverlapArea(LayoutBounds first, LayoutBounds second) {
            double width = Math.Min(first.Right, second.Right) - Math.Max(first.Left, second.Left);
            double height = Math.Min(first.Top, second.Top) - Math.Max(first.Bottom, second.Bottom);
            return width > 0D && height > 0D ? width * height : 0D;
        }

        private List<LayoutBounds> BuildNoteReservedBounds(VisioPage page) {
            List<LayoutBounds> reservedBounds = new();
            foreach (VisioShape shape in page.Shapes) {
                if (IsIgnoredNoteObstacle(shape)) {
                    continue;
                }

                reservedBounds.Add(LayoutBounds.FromCenter(shape.PinX, shape.PinY, shape.Width, shape.Height).Inflate(NoteCollisionPadding));
            }

            foreach (VisioConnector connector in page.Connectors) {
                AddConnectorSegmentReservedBounds(connector, reservedBounds);
                VisioConnectorLabelPlacement? label = connector.LabelPlacement;
                if (label != null && label.PinX.HasValue && label.PinY.HasValue) {
                    reservedBounds.Add(LayoutBounds.FromCenter(label.PinX.Value, label.PinY.Value, label.Width, label.Height).Inflate(NoteCollisionPadding));
                }
            }

            return reservedBounds;
        }

        private static void AddConnectorSegmentReservedBounds(VisioConnector connector, IList<LayoutBounds> reservedBounds) {
            List<(double X, double Y)> points = new() {
                (connector.From.PinX, connector.From.PinY)
            };
            foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                points.Add((waypoint.X, waypoint.Y));
            }

            if (connector.Waypoints.Count == 0 && connector.Kind == ConnectorKind.RightAngle) {
                points.Add((connector.From.PinX, connector.To.PinY));
            }

            points.Add((connector.To.PinX, connector.To.PinY));
            const double thickness = 0.08D;
            for (int i = 0; i < points.Count - 1; i++) {
                double left = Math.Min(points[i].X, points[i + 1].X) - (thickness / 2D);
                double right = Math.Max(points[i].X, points[i + 1].X) + (thickness / 2D);
                double bottom = Math.Min(points[i].Y, points[i + 1].Y) - (thickness / 2D);
                double top = Math.Max(points[i].Y, points[i + 1].Y) + (thickness / 2D);
                reservedBounds.Add(new LayoutBounds(left, top, right, bottom));
            }
        }

        private static bool IsIgnoredNoteObstacle(VisioShape shape) {
            if (shape.Width <= 0.06D && shape.Height <= 0.06D) {
                return true;
            }

            string? kind = shape.GetUserCellValue(VisioSemanticUserCells.Kind);
            return string.Equals(kind, VisioSemanticUserCells.SequenceFragmentKind, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(kind, VisioSemanticUserCells.BackgroundSurfaceKind, StringComparison.OrdinalIgnoreCase);
        }
    }
}
