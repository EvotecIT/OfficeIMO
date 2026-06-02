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
        private const string LifelineLayer = "Sequence Lifelines";
        private const string MessageLayer = "Sequence Messages";
        private const string ActivationLayer = "Sequence Activations";
        private const string FragmentLayer = "Sequence Fragments";
        private const string NoteLayer = "Sequence Notes";
        private const string GuideLayer = "Sequence Guides";

        private sealed class ParticipantItem {
            public ParticipantItem(string id, string text, VisioSequenceParticipantKind kind) {
                Id = id;
                Text = text;
                Kind = kind;
            }

            public string Id { get; }

            public string Text { get; }

            public VisioSequenceParticipantKind Kind { get; }

            public double PinX { get; set; }

            public VisioShape? Header { get; set; }

            public VisioShape? BottomAnchor { get; set; }
        }

        private sealed class MessageItem {
            public MessageItem(string id, string fromId, string toId, string label, VisioSequenceMessageKind kind, bool selfMessage) {
                Id = id;
                FromId = fromId;
                ToId = toId;
                Label = label;
                Kind = kind;
                SelfMessage = selfMessage;
            }

            public string Id { get; }

            public string FromId { get; }

            public string ToId { get; }

            public string Label { get; }

            public VisioSequenceMessageKind Kind { get; }

            public bool SelfMessage { get; }
        }

        private sealed class NoteItem {
            public NoteItem(string id, string participantId, string text, int rowIndex, VisioSide placement) {
                Id = id;
                ParticipantId = participantId;
                Text = text;
                RowIndex = rowIndex;
                Placement = placement;
            }

            public string Id { get; }

            public string ParticipantId { get; }

            public string Text { get; }

            public int RowIndex { get; }

            public VisioSide Placement { get; }
        }

        private sealed class ActivationItem {
            public ActivationItem(string id, string participantId, int startRowIndex, int endRowIndex) {
                Id = id;
                ParticipantId = participantId;
                StartRowIndex = startRowIndex;
                EndRowIndex = endRowIndex;
            }

            public string Id { get; }

            public string ParticipantId { get; }

            public int StartRowIndex { get; }

            public int EndRowIndex { get; }
        }

        private sealed class FragmentItem {
            public FragmentItem(string id, string text, int startRowIndex, int endRowIndex, IReadOnlyList<string> participantIds, string? parentFragmentId) {
                Id = id;
                Text = text;
                StartRowIndex = startRowIndex;
                EndRowIndex = endRowIndex;
                ParticipantIds = participantIds;
                ParentFragmentId = parentFragmentId;
            }

            public string Id { get; }

            public string Text { get; }

            public int StartRowIndex { get; }

            public int EndRowIndex { get; }

            public IReadOnlyList<string> ParticipantIds { get; }

            public string? ParentFragmentId { get; }
        }

        private sealed class FragmentOperandItem {
            public FragmentOperandItem(string id, string fragmentId, string text, int rowIndex, bool divider) {
                Id = id;
                FragmentId = fragmentId;
                Text = text;
                RowIndex = rowIndex;
                Divider = divider;
            }

            public string Id { get; }

            public string FragmentId { get; }

            public string Text { get; }

            public int RowIndex { get; }

            public bool Divider { get; }
        }

        private readonly struct LayoutBounds {
            public LayoutBounds(double left, double top, double right, double bottom) {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            public double Left { get; }

            public double Top { get; }

            public double Right { get; }

            public double Bottom { get; }

            public static LayoutBounds FromCenter(double x, double y, double width, double height) {
                double halfWidth = width / 2D;
                double halfHeight = height / 2D;
                return new LayoutBounds(x - halfWidth, y + halfHeight, x + halfWidth, y - halfHeight);
            }

            public LayoutBounds Inflate(double padding) =>
                new(Left - padding, Top + padding, Right + padding, Bottom - padding);
        }

        private readonly struct NotePlacement {
            public NotePlacement(double x, double y, VisioSide resolvedPlacement, LayoutBounds bounds) {
                X = x;
                Y = y;
                ResolvedPlacement = resolvedPlacement;
                Bounds = bounds;
            }

            public double X { get; }

            public double Y { get; }

            public VisioSide ResolvedPlacement { get; }

            public LayoutBounds Bounds { get; }
        }

        private readonly struct MessageLabelPlacement {
            public MessageLabelPlacement(double x, double y, double width, double height) {
                X = x;
                Y = y;
                Width = width;
                Height = height;
            }

            public double X { get; }

            public double Y { get; }

            public double Width { get; }

            public double Height { get; }
        }

        private readonly struct FragmentLayout {
            public FragmentLayout(double left, double right, double topY, double bottomY, double width, double height, double x, double y, double labelLeft, double labelWidth, double labelY, double firstParticipantX, int depth, int overlapLane) {
                Left = left;
                Right = right;
                TopY = topY;
                BottomY = bottomY;
                Width = width;
                Height = height;
                X = x;
                Y = y;
                LabelLeft = labelLeft;
                LabelWidth = labelWidth;
                LabelY = labelY;
                FirstParticipantX = firstParticipantX;
                Depth = depth;
                OverlapLane = overlapLane;
            }

            public double Left { get; }

            public double Right { get; }

            public double TopY { get; }

            public double BottomY { get; }

            public double Width { get; }

            public double Height { get; }

            public double X { get; }

            public double Y { get; }

            public double LabelLeft { get; }

            public double LabelWidth { get; }

            public double LabelY { get; }

            public double FirstParticipantX { get; }

            public int Depth { get; }

            public int OverlapLane { get; }
        }
    }
}
