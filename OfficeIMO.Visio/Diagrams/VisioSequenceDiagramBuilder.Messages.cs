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
        private void AddMessages(VisioPage page, double firstMessageY) {
            for (int i = 0; i < _messages.Count; i++) {
                MessageItem message = _messages[i];
                double y = firstMessageY - (i * _messageSpacing);
                if (message.SelfMessage) {
                    AddSelfMessage(page, message, y);
                } else {
                    AddParticipantMessage(page, message, y, firstMessageY);
                }
            }
        }

        private void AddActivations(VisioPage page, double firstMessageY) {
            foreach (ActivationItem activation in _activations) {
                ParticipantItem participant = _participantsById[activation.ParticipantId];
                double topY = firstMessageY - (activation.StartRowIndex * _messageSpacing) + (_messageSpacing * 0.32D);
                double bottomY = firstMessageY - (activation.EndRowIndex * _messageSpacing) - (_messageSpacing * 0.32D);
                double height = Math.Max(ActivationMinimumHeight, topY - bottomY);
                double y = (topY + bottomY) / 2D;
                VisioShape shape = page.AddStencilShape(VisioStencils.Sequence, "seq.activation", activation.Id, participant.PinX, y, ActivationWidth, height, string.Empty);
                _theme.Marker.ApplyTo(shape);
                shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.SequenceActivationKind, "STR", prompt: "OfficeIMO semantic kind");
                shape.SetUserCell("OfficeIMO.SequenceParticipantId", activation.ParticipantId, "STR", prompt: "OfficeIMO sequence activation participant");
                shape.SetUserCell("OfficeIMO.SequenceStartRowIndex", activation.StartRowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence activation start row");
                shape.SetUserCell("OfficeIMO.SequenceEndRowIndex", activation.EndRowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence activation end row");
                page.AddToLayer(ActivationLayer, shape);
            }
        }


        private void AddParticipantMessage(VisioPage page, MessageItem message, double y, double firstMessageY) {
            ParticipantItem from = _participantsById[message.FromId];
            ParticipantItem to = _participantsById[message.ToId];
            bool leftToRight = from.PinX <= to.PinX;
            VisioShape fromAnchor = CreateAnchor(page, message.Id + "-from", from.PinX, y);
            VisioShape toAnchor = CreateAnchor(page, message.Id + "-to", to.PinX, y);
            VisioConnector connector = page.AddConnector(
                message.Id,
                fromAnchor,
                toAnchor,
                ConnectorKind.Straight,
                leftToRight ? VisioSide.Right : VisioSide.Left,
                leftToRight ? VisioSide.Left : VisioSide.Right);

            ApplyMessageStyle(connector, message.Kind);
            connector.Label = message.Label;
            MessageLabelPlacement label = ResolveParticipantMessageLabelPlacement(page, from, to, message, y, firstMessageY);
            connector.PlaceLabelAt(label.X, label.Y, label.Width, label.Height);
            page.AddToLayer(MessageLayer, connector);
        }

        private MessageLabelPlacement ResolveParticipantMessageLabelPlacement(VisioPage page, ParticipantItem from, ParticipantItem to, MessageItem message, double y, double firstMessageY) {
            double left = Math.Min(from.PinX, to.PinX);
            double right = Math.Max(from.PinX, to.PinX);
            double span = Math.Max(0D, right - left);
            double desiredWidth = EstimateParticipantMessageLabelWidth(message.Label, span);
            double labelY = y + MessageLabelVerticalOffset;
            double midpoint = (from.PinX + to.PinX) / 2D;
            MessageLabelPlacement best = new(midpoint, labelY, desiredWidth, MessageLabelHeight);
            double bestScore = ScoreParticipantMessageLabel(page, best, from, to, y, firstMessageY, midpoint);

            foreach (double candidateX in GetParticipantMessageLabelCandidateCenters(left, right)) {
                double availableWidth = GetParticipantMessageLabelAvailableWidth(candidateX, left, right);
                if (availableWidth < 0.65D) {
                    continue;
                }

                double width = Math.Max(0.9D, Math.Min(desiredWidth, availableWidth));
                MessageLabelPlacement candidate = new(candidateX, labelY, width, MessageLabelHeight);
                double score = ScoreParticipantMessageLabel(page, candidate, from, to, y, firstMessageY, midpoint);
                if (score < bestScore) {
                    best = candidate;
                    bestScore = score;
                }
            }

            return best;
        }

        private IEnumerable<double> GetParticipantMessageLabelCandidateCenters(double left, double right) {
            yield return (left + right) / 2D;

            List<double> anchors = new() { left, right };
            anchors.AddRange(_participants
                .Select(participant => participant.PinX)
                .Where(x => x > left && x < right));
            anchors = anchors
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            for (int i = 0; i < anchors.Count - 1; i++) {
                double gapLeft = anchors[i] + MessageLabelLifelinePadding + (ActivationWidth / 2D);
                double gapRight = anchors[i + 1] - MessageLabelLifelinePadding - (ActivationWidth / 2D);
                if (gapRight > gapLeft) {
                    yield return (gapLeft + gapRight) / 2D;
                }
            }
        }

        private double GetParticipantMessageLabelAvailableWidth(double candidateX, double left, double right) {
            double nearestLeft = left;
            double nearestRight = right;
            foreach (ParticipantItem participant in _participants) {
                if (participant.PinX <= left || participant.PinX >= right) {
                    continue;
                }

                if (participant.PinX < candidateX) {
                    nearestLeft = Math.Max(nearestLeft, participant.PinX);
                } else if (participant.PinX > candidateX) {
                    nearestRight = Math.Min(nearestRight, participant.PinX);
                }
            }

            double available = nearestRight - nearestLeft - (2D * MessageLabelLifelinePadding) - ActivationWidth;
            return Math.Max(0D, available);
        }

        private double ScoreParticipantMessageLabel(VisioPage page, MessageLabelPlacement placement, ParticipantItem from, ParticipantItem to, double messageY, double firstMessageY, double midpoint) {
            LayoutBounds bounds = LayoutBounds.FromCenter(placement.X, placement.Y, placement.Width, placement.Height);
            double left = Math.Min(from.PinX, to.PinX);
            double right = Math.Max(from.PinX, to.PinX);
            double score = Math.Abs(placement.X - midpoint) * 0.4D;

            foreach (ParticipantItem participant in _participants) {
                if (participant.PinX < left || participant.PinX > right) {
                    continue;
                }

                LayoutBounds lifeline = new(
                    participant.PinX - MessageLabelLifelinePadding,
                    bounds.Top,
                    participant.PinX + MessageLabelLifelinePadding,
                    bounds.Bottom);
                score += GetOverlapArea(bounds, lifeline) * 35D;
            }

            foreach (ActivationItem activation in _activations) {
                ParticipantItem participant = _participantsById[activation.ParticipantId];
                if (participant.PinX < left || participant.PinX > right || !IsActivationNearMessageRow(activation, messageY)) {
                    continue;
                }

                LayoutBounds activationBounds = new(
                    participant.PinX - (ActivationWidth / 2D) - MessageLabelActivationPadding,
                    bounds.Top,
                    participant.PinX + (ActivationWidth / 2D) + MessageLabelActivationPadding,
                    bounds.Bottom);
                score += GetOverlapArea(bounds, activationBounds) * 45D;
            }

            foreach (LayoutBounds obstacle in GetFragmentLabelObstacleBounds(page, firstMessageY)) {
                score += GetOverlapArea(bounds.Inflate(0.03D), obstacle.Inflate(0.04D)) * 70D;
            }

            return score;
        }

        private IEnumerable<LayoutBounds> GetFragmentLabelObstacleBounds(VisioPage page, double firstMessageY) {
            foreach (FragmentItem fragment in _fragments) {
                FragmentLayout layout = CalculateFragmentLayout(page, fragment, firstMessageY);
                yield return LayoutBounds.FromCenter(layout.LabelLeft + (layout.LabelWidth / 2D), layout.LabelY, layout.LabelWidth, FragmentHeaderHeight);

                foreach (FragmentOperandItem operand in _fragmentOperands.Where(item => string.Equals(item.FragmentId, fragment.Id, StringComparison.Ordinal)).OrderBy(item => item.RowIndex).ThenBy(item => item.Id, StringComparer.Ordinal)) {
                    double rowY = firstMessageY - (operand.RowIndex * _messageSpacing);
                    double dividerY = rowY + (_messageSpacing * 0.46D);
                    double operandLabelY = operand.Divider
                        ? dividerY - 0.07D - (FragmentOperandLabelHeight / 2D)
                        : Math.Min(layout.TopY - FragmentHeaderHeight - (FragmentOperandLabelHeight / 2D) - 0.06D, rowY + 0.2D);
                    double operandLabelWidth = Math.Max(0.72D, Math.Min(layout.Width - 0.16D, 0.42D + (operand.Text.Trim().Length * 0.055D)));
                    double operandLabelLeft = Math.Max(layout.Left + 0.12D, layout.FirstParticipantX + (ActivationWidth / 2D) + 0.12D);
                    operandLabelLeft = Math.Min(operandLabelLeft, layout.Left + layout.Width - operandLabelWidth - 0.04D);
                    yield return LayoutBounds.FromCenter(operandLabelLeft + (operandLabelWidth / 2D), operandLabelY, operandLabelWidth, FragmentOperandLabelHeight);
                }
            }
        }

        private bool IsActivationNearMessageRow(ActivationItem activation, double messageY) {
            double firstMessageY = _participants.Count == 0 || _participants[0].Header == null
                ? messageY
                : _participants[0].Header!.PinY - (_participantHeight / 2D) - _messageGap;
            double activationTop = firstMessageY - (activation.StartRowIndex * _messageSpacing) + (_messageSpacing * 0.32D);
            double activationBottom = firstMessageY - (activation.EndRowIndex * _messageSpacing) - (_messageSpacing * 0.32D);
            return messageY <= activationTop && messageY >= activationBottom;
        }

        private static double EstimateParticipantMessageLabelWidth(string label, double span) {
            double maxWidth = Math.Max(0.9D, Math.Min(2.6D, span - 0.2D));
            if (string.IsNullOrWhiteSpace(label)) {
                return Math.Min(1.2D, maxWidth);
            }

            double estimatedWidth = 0.48D + (label.Trim().Length * 0.065D);
            return Math.Max(0.9D, Math.Min(maxWidth, estimatedWidth));
        }

        private void AddSelfMessage(VisioPage page, MessageItem message, double y) {
            ParticipantItem participant = _participantsById[message.FromId];
            double lowerY = y - Math.Min(_selfMessageHeight, _messageSpacing * 0.55D);
            VisioShape fromAnchor = CreateAnchor(page, message.Id + "-from", participant.PinX, y);
            VisioShape toAnchor = CreateAnchor(page, message.Id + "-to", participant.PinX, lowerY);
            ResolveSelfMessageLabelPlacement(page, participant, message.Label, out double direction, out double labelWidth, out double labelHeight);
            VisioSide connectorSide = direction > 0D ? VisioSide.Right : VisioSide.Left;
            double loopX = participant.PinX + (direction * _selfMessageWidth);
            VisioConnector connector = page.AddConnector(message.Id, fromAnchor, toAnchor, ConnectorKind.RightAngle, connectorSide, connectorSide);
            ApplyMessageStyle(connector, message.Kind);
            connector.Label = message.Label;
            connector.RouteThrough(
                VisioConnectorWaypoint.At(loopX, y),
                VisioConnectorWaypoint.At(loopX, lowerY));
            double labelCenterX = loopX + (direction * (SelfMessageLabelGap + (labelWidth / 2D)));
            double labelCenterY = y - (labelHeight / 2D);
            connector.PlaceLabelAt(labelCenterX, labelCenterY, labelWidth, labelHeight);
            page.AddToLayer(MessageLayer, connector);
        }


        private void ResolveSelfMessageLabelPlacement(VisioPage page, ParticipantItem participant, string label, out double direction, out double labelWidth, out double labelHeight) {
            double desiredWidth = EstimateSelfMessageLabelWidth(label);
            double rightAvailable = GetSelfMessageLabelAvailableWidth(page, participant, rightSide: true);
            double leftAvailable = GetSelfMessageLabelAvailableWidth(page, participant, rightSide: false);
            bool useRight = rightAvailable >= Math.Min(desiredWidth, 1.1D) || rightAvailable >= leftAvailable;
            double available = Math.Max(0D, useRight ? rightAvailable : leftAvailable);

            direction = useRight ? 1D : -1D;
            labelWidth = Math.Max(0.2D, Math.Min(desiredWidth, available));
            labelHeight = desiredWidth > labelWidth + 0.05D ? 0.46D : SelfMessageLabelHeight;
        }

        private double GetSelfMessageLabelAvailableWidth(VisioPage page, ParticipantItem participant, bool rightSide) {
            double direction = rightSide ? 1D : -1D;
            double loopX = participant.PinX + (direction * _selfMessageWidth);
            double pageWidth = page.Width.FromInches(_unit);
            double pageLimit = rightSide
                ? pageWidth - _rightMargin - loopX - SelfMessageLabelGap
                : loopX - _leftMargin - SelfMessageLabelGap;
            double nearestParticipantLimit = double.PositiveInfinity;

            foreach (ParticipantItem other in _participants) {
                if (ReferenceEquals(other, participant)) {
                    continue;
                }

                if (rightSide && other.PinX > participant.PinX) {
                    nearestParticipantLimit = Math.Min(nearestParticipantLimit, other.PinX - (_participantWidth / 2D) - loopX - SelfMessageLabelGap - 0.18D);
                } else if (!rightSide && other.PinX < participant.PinX) {
                    nearestParticipantLimit = Math.Min(nearestParticipantLimit, loopX - (other.PinX + (_participantWidth / 2D)) - SelfMessageLabelGap - 0.18D);
                }
            }

            return Math.Max(0D, Math.Min(pageLimit, nearestParticipantLimit));
        }

        private static double EstimateSelfMessageLabelWidth(string label) {
            if (string.IsNullOrWhiteSpace(label)) {
                return 1.2D;
            }

            double estimatedWidth = 0.55D + (label.Trim().Length * 0.075D);
            return Math.Max(1.2D, Math.Min(2.4D, estimatedWidth));
        }
    }
}
