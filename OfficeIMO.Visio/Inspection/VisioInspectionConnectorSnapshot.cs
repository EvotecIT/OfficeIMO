using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Visio {
/// <summary>
    /// Snapshot of a Visio connector.
    /// </summary>
    public sealed class VisioInspectionConnectorSnapshot {
        internal VisioInspectionConnectorSnapshot(
            string id,
            string fromId,
            string toId,
            string kind,
            string? label,
            bool hasLabelPlacement,
            double? labelPosition,
            double? labelOffsetX,
            double? labelOffsetY,
            double? labelPinX,
            double? labelPinY,
            double? labelResolvedPinX,
            double? labelResolvedPinY,
            double? labelLocPinX,
            double? labelLocPinY,
            double? labelWidth,
            double? labelHeight,
            IReadOnlyList<VisioInspectionWaypointSnapshot> waypoints,
            string lineColor,
            int linePattern,
            double lineWeight,
            string? beginArrow,
            string? endArrow,
            IReadOnlyList<string> layers,
            IReadOnlyList<VisioInspectionShapeDataSnapshot> shapeData,
            IReadOnlyDictionary<string, string> data) {
            Id = id;
            FromId = fromId;
            ToId = toId;
            Kind = kind;
            Label = label;
            HasLabelPlacement = hasLabelPlacement;
            LabelPosition = labelPosition;
            LabelOffsetX = labelOffsetX;
            LabelOffsetY = labelOffsetY;
            LabelPinX = labelPinX;
            LabelPinY = labelPinY;
            LabelResolvedPinX = labelResolvedPinX;
            LabelResolvedPinY = labelResolvedPinY;
            LabelLocPinX = labelLocPinX;
            LabelLocPinY = labelLocPinY;
            LabelWidth = labelWidth;
            LabelHeight = labelHeight;
            Waypoints = waypoints;
            LineColor = lineColor;
            LinePattern = linePattern;
            LineWeight = lineWeight;
            BeginArrow = beginArrow;
            EndArrow = endArrow;
            Layers = layers;
            ShapeData = shapeData;
            Data = data;
        }

        /// <summary>Connector identifier.</summary>
        public string Id { get; }

        /// <summary>Source shape identifier.</summary>
        public string FromId { get; }

        /// <summary>Target shape identifier.</summary>
        public string ToId { get; }

        /// <summary>Connector kind.</summary>
        public string Kind { get; }

        /// <summary>Connector label.</summary>
        public string? Label { get; }

        /// <summary>Whether explicit label placement exists.</summary>
        public bool HasLabelPlacement { get; }

        /// <summary>Relative label position along the connector path, when explicit placement exists.</summary>
        public double? LabelPosition { get; }

        /// <summary>Relative label X offset, when explicit placement exists.</summary>
        public double? LabelOffsetX { get; }

        /// <summary>Relative label Y offset, when explicit placement exists.</summary>
        public double? LabelOffsetY { get; }

        /// <summary>Absolute label X coordinate, when the label is pinned to the page.</summary>
        public double? LabelPinX { get; }

        /// <summary>Absolute label Y coordinate, when the label is pinned to the page.</summary>
        public double? LabelPinY { get; }

        /// <summary>Resolved page X coordinate for the label pin.</summary>
        public double? LabelResolvedPinX { get; }

        /// <summary>Resolved page Y coordinate for the label pin.</summary>
        public double? LabelResolvedPinY { get; }

        /// <summary>Resolved local X pin inside the label text box.</summary>
        public double? LabelLocPinX { get; }

        /// <summary>Resolved local Y pin inside the label text box.</summary>
        public double? LabelLocPinY { get; }

        /// <summary>Explicit label width, when explicit placement exists.</summary>
        public double? LabelWidth { get; }

        /// <summary>Explicit label height, when explicit placement exists.</summary>
        public double? LabelHeight { get; }

        /// <summary>Explicit connector waypoints.</summary>
        public IReadOnlyList<VisioInspectionWaypointSnapshot> Waypoints { get; }

        /// <summary>Line color as a stable OfficeIMO color string.</summary>
        public string LineColor { get; }

        /// <summary>Visio line pattern value.</summary>
        public int LinePattern { get; }

        /// <summary>Connector line weight.</summary>
        public double LineWeight { get; }

        /// <summary>Begin arrow value.</summary>
        public string? BeginArrow { get; }

        /// <summary>End arrow value.</summary>
        public string? EndArrow { get; }

        /// <summary>Layer names assigned to the connector.</summary>
        public IReadOnlyList<string> Layers { get; }

        /// <summary>Shape Data rows attached to the connector.</summary>
        public IReadOnlyList<VisioInspectionShapeDataSnapshot> ShapeData { get; }

        /// <summary>Arbitrary data attached to the connector.</summary>
        public IReadOnlyDictionary<string, string> Data { get; }

        internal void AppendText(StringBuilder builder, string pagePrefix) {
            string prefix = pagePrefix + ".connector[" + VisioInspectionSnapshot.EscapeKey(Id) + "]";
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".from", FromId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".to", ToId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".kind", Kind);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".label", Label);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".hasLabelPlacement", HasLabelPlacement);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelPosition", LabelPosition);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelOffsetX", LabelOffsetX);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelOffsetY", LabelOffsetY);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelPinX", LabelPinX);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelPinY", LabelPinY);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelResolvedPinX", LabelResolvedPinX);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelResolvedPinY", LabelResolvedPinY);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelLocPinX", LabelLocPinX);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelLocPinY", LabelLocPinY);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelWidth", LabelWidth);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".labelHeight", LabelHeight);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".lineColor", LineColor);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".linePattern", LinePattern);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".lineWeight", LineWeight);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".beginArrow", BeginArrow);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".endArrow", EndArrow);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".layers", string.Join(",", Layers));

            for (int i = 0; i < Waypoints.Count; i++) {
                string waypointPrefix = prefix + ".waypoint[" + i.ToString(CultureInfo.InvariantCulture) + "]";
                VisioInspectionSnapshot.AppendLine(builder, waypointPrefix + ".x", Waypoints[i].X);
                VisioInspectionSnapshot.AppendLine(builder, waypointPrefix + ".y", Waypoints[i].Y);
            }

            VisioInspectionShapeSnapshot.AppendShapeData(builder, prefix, ShapeData);
            VisioInspectionShapeSnapshot.AppendData(builder, prefix, Data);
        }
    }
}
