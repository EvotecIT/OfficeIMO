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
        private void AddFragments(VisioPage page, double firstMessageY) {
            foreach (FragmentItem fragment in _fragments.OrderBy(GetFragmentDepth).ThenBy(item => item.StartRowIndex).ThenByDescending(item => item.EndRowIndex).ThenBy(item => item.Id, StringComparer.Ordinal)) {
                FragmentLayout layout = CalculateFragmentLayout(page, fragment, firstMessageY);
                VisioShape frame = page.AddStencilShape(VisioStencils.Sequence, "seq.fragment", fragment.Id, layout.X, layout.Y, layout.Width, layout.Height, string.Empty);
                frame.FillColor = OfficeColor.Transparent;
                frame.FillPattern = 0;
                frame.LineColor = _theme.ControlConnector.LineColor;
                frame.LinePattern = 2;
                frame.LineWeight = Math.Max(0.012D, _theme.ControlConnector.LineWeight);
                frame.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.SequenceFragmentKind, "STR", prompt: "OfficeIMO semantic kind");
                frame.SetUserCell("OfficeIMO.SequenceStartRowIndex", fragment.StartRowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence fragment start row");
                frame.SetUserCell("OfficeIMO.SequenceEndRowIndex", fragment.EndRowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence fragment end row");
                frame.SetUserCell("OfficeIMO.SequenceParticipantIds", string.Join(";", fragment.ParticipantIds), "STR", prompt: "OfficeIMO sequence fragment participants");
                if (!string.IsNullOrWhiteSpace(fragment.ParentFragmentId)) {
                    frame.SetUserCell("OfficeIMO.SequenceParentFragmentId", fragment.ParentFragmentId, "STR", prompt: "OfficeIMO parent sequence fragment");
                }

                frame.SetUserCell("OfficeIMO.SequenceFragmentDepth", layout.Depth.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence fragment nesting depth");
                frame.SetUserCell("OfficeIMO.SequenceFragmentOverlapLane", layout.OverlapLane.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence fragment overlap lane");
                page.AddToLayer(FragmentLayer, frame);

                VisioShape label = page.AddTextBox(GetFragmentLabelId(fragment.Id), layout.LabelLeft + (layout.LabelWidth / 2D), layout.LabelY, layout.LabelWidth, FragmentHeaderHeight, fragment.Text, _unit);
                label.FillColor = _theme.Container.FillColor;
                label.FillPattern = 1;
                label.LineColor = OfficeColor.Transparent;
                label.LinePattern = 0;
                label.TextStyle = CreateFragmentLabelTextStyle();
                VisioSemanticUserCells.MarkGeneratedAdornment(label);
                label.SetUserCell("OfficeIMO.SequenceFragmentId", fragment.Id, "STR", prompt: "OfficeIMO sequence fragment label target");
                page.AddToLayer(FragmentLayer, label);

                foreach (FragmentOperandItem operand in _fragmentOperands.Where(item => string.Equals(item.FragmentId, fragment.Id, StringComparison.Ordinal)).OrderBy(item => item.RowIndex).ThenBy(item => item.Id, StringComparer.Ordinal)) {
                    AddFragmentOperand(page, fragment, operand, layout, firstMessageY);
                }
            }
        }

        private FragmentLayout CalculateFragmentLayout(VisioPage page, FragmentItem fragment, double firstMessageY) {
            IReadOnlyList<ParticipantItem> participants = fragment.ParticipantIds
                .Select(id => _participantsById[id])
                .ToArray();
            double left = participants.Min(participant => participant.PinX) - (_participantWidth / 2D) - FragmentHorizontalPadding;
            double right = participants.Max(participant => participant.PinX) + (_participantWidth / 2D) + FragmentHorizontalPadding;
            left = Math.Max(_leftMargin, left);
            double pageWidth = page.Width.FromInches(_unit);
            right = Math.Min(pageWidth - _rightMargin, right);

            double topY = firstMessageY - (fragment.StartRowIndex * _messageSpacing) + (_messageSpacing * 0.46D) + FragmentVerticalPadding;
            double headerBottomY = firstMessageY + _messageGap;
            topY = Math.Min(topY, headerBottomY - 0.06D);
            double bottomY = firstMessageY - (fragment.EndRowIndex * _messageSpacing) - (_messageSpacing * 0.46D) - FragmentVerticalPadding;
            int depth = GetFragmentDepth(fragment);
            int overlapLane = GetFragmentOverlapLane(fragment);
            double laneInset = (depth + overlapLane) * FragmentNestedHorizontalInset;
            double verticalInset = (depth + overlapLane) * FragmentNestedVerticalInset;

            if (!string.IsNullOrWhiteSpace(fragment.ParentFragmentId)) {
                FragmentLayout parentLayout = CalculateFragmentLayout(page, RequireFragment(fragment.ParentFragmentId!), firstMessageY);
                left = Math.Max(left, parentLayout.Left + FragmentNestedHorizontalInset);
                right = Math.Min(right, parentLayout.Right - FragmentNestedHorizontalInset);
                topY = Math.Min(topY, parentLayout.TopY - FragmentHeaderHeight - FragmentNestedVerticalInset);
                bottomY = Math.Max(bottomY, parentLayout.BottomY + FragmentNestedVerticalInset);
            }

            left += laneInset;
            right -= laneInset;
            topY -= verticalInset;
            bottomY += verticalInset;
            if (right <= left) {
                double center = (left + right) / 2D;
                left = center - (FragmentMinimumWidth / 2D);
                right = center + (FragmentMinimumWidth / 2D);
            }

            if (topY <= bottomY) {
                double center = (topY + bottomY) / 2D;
                topY = center + (FragmentMinimumHeight / 2D);
                bottomY = center - (FragmentMinimumHeight / 2D);
            }

            double width = Math.Max(FragmentMinimumWidth, right - left);
            double height = Math.Max(FragmentMinimumHeight, topY - bottomY);
            double x = left + (width / 2D);
            double y = (topY + bottomY) / 2D;
            double labelWidth = Math.Max(0.75D, Math.Min(width - 0.1D, 0.48D + (fragment.Text.Trim().Length * 0.07D)));
            double labelLeft = left + 0.04D + (overlapLane * 0.08D);
            double firstParticipantX = participants.Min(participant => participant.PinX);
            if (labelLeft < firstParticipantX + ActivationWidth && labelLeft + labelWidth > firstParticipantX - ActivationWidth) {
                labelLeft = firstParticipantX + (ActivationWidth / 2D) + 0.08D;
                labelLeft = Math.Min(labelLeft, left + width - labelWidth - 0.04D);
            }

            double labelY = topY - 0.04D - (FragmentHeaderHeight / 2D);
            return new FragmentLayout(left, right, topY, bottomY, width, height, x, y, labelLeft, labelWidth, labelY, firstParticipantX, depth, overlapLane);
        }

        private int GetFragmentDepth(FragmentItem fragment) {
            int depth = 0;
            string? parentId = fragment.ParentFragmentId;
            while (!string.IsNullOrWhiteSpace(parentId)) {
                depth++;
                FragmentItem parent = RequireFragment(parentId!);
                parentId = parent.ParentFragmentId;
            }

            return depth;
        }

        private int GetFragmentOverlapLane(FragmentItem fragment) {
            int lane = 0;
            foreach (FragmentItem other in _fragments) {
                if (ReferenceEquals(other, fragment)) {
                    break;
                }

                if (GetFragmentDepth(other) == GetFragmentDepth(fragment) &&
                    string.Equals(other.ParentFragmentId, fragment.ParentFragmentId, StringComparison.Ordinal) &&
                    FragmentRowsOverlap(other, fragment) &&
                    FragmentParticipantsOverlap(other, fragment)) {
                    lane++;
                }
            }

            return lane;
        }

        private static bool FragmentRowsOverlap(FragmentItem first, FragmentItem second) =>
            first.StartRowIndex <= second.EndRowIndex && second.StartRowIndex <= first.EndRowIndex;

        private static bool FragmentParticipantsOverlap(FragmentItem first, FragmentItem second) =>
            first.ParticipantIds.Any(id => second.ParticipantIds.Contains(id, StringComparer.Ordinal));

        private void AddFragmentOperand(VisioPage page, FragmentItem fragment, FragmentOperandItem operand, FragmentLayout layout, double firstMessageY) {
            double rowY = firstMessageY - (operand.RowIndex * _messageSpacing);
            double dividerY = rowY + (_messageSpacing * 0.46D);
            double labelY = operand.Divider
                ? dividerY - 0.07D - (FragmentOperandLabelHeight / 2D)
                : Math.Min(layout.TopY - FragmentHeaderHeight - (FragmentOperandLabelHeight / 2D) - 0.06D, rowY + 0.2D);

            if (operand.Divider) {
                VisioShape leftAnchor = CreateAnchor(page, operand.Id + "-from", layout.Left + 0.04D, dividerY);
                VisioShape rightAnchor = CreateAnchor(page, operand.Id + "-to", layout.Right - 0.04D, dividerY);
                VisioConnector divider = page.AddConnector(operand.Id, leftAnchor, rightAnchor, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
                divider.LineColor = _theme.ControlConnector.LineColor;
                divider.LinePattern = 2;
                divider.LineWeight = Math.Max(0.008D, _theme.ControlConnector.LineWeight * 0.75D);
                divider.EndArrow = EndArrow.None;
                divider.BeginArrow = EndArrow.None;
                page.AddToLayer(FragmentLayer, divider);
            }

            double labelWidth = Math.Max(0.72D, Math.Min(layout.Width - 0.16D, 0.42D + (operand.Text.Trim().Length * 0.055D)));
            double labelLeft = Math.Max(layout.Left + 0.12D, layout.FirstParticipantX + (ActivationWidth / 2D) + 0.12D);
            labelLeft = Math.Min(labelLeft, layout.Left + layout.Width - labelWidth - 0.04D);
            VisioShape label = page.AddTextBox(GetFragmentOperandLabelId(operand.Id), labelLeft + (labelWidth / 2D), labelY, labelWidth, FragmentOperandLabelHeight, operand.Text, _unit);
            label.FillColor = _theme.Container.FillColor;
            label.FillPattern = 1;
            label.LineColor = OfficeColor.Transparent;
            label.LinePattern = 0;
            label.TextStyle = CreateFragmentOperandTextStyle();
            VisioSemanticUserCells.MarkGeneratedAdornment(label);
            label.SetUserCell("OfficeIMO.SequenceFragmentId", fragment.Id, "STR", prompt: "OfficeIMO sequence fragment label target");
            label.SetUserCell("OfficeIMO.SequenceFragmentOperandId", operand.Id, "STR", prompt: "OfficeIMO sequence fragment operand id");
            label.SetUserCell("OfficeIMO.SequenceFragmentOperandRowIndex", operand.RowIndex.ToString(global::System.Globalization.CultureInfo.InvariantCulture), "STR", prompt: "OfficeIMO sequence fragment operand row");
            label.SetUserCell("OfficeIMO.SequenceFragmentOperandDivider", operand.Divider ? "true" : "false", "STR", prompt: "OfficeIMO sequence fragment operand divider");
            page.AddToLayer(FragmentLayer, label);
        }
    }
}
