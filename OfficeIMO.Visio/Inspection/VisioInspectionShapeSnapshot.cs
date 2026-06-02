using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
/// <summary>
    /// Snapshot of a Visio shape.
    /// </summary>
    public sealed class VisioInspectionShapeSnapshot {
        internal VisioInspectionShapeSnapshot(
            string id,
            string? name,
            string? nameU,
            string? type,
            string? masterId,
            string? masterNameU,
            string? masterShapeId,
            string? parentId,
            string? text,
            double pinX,
            double pinY,
            double width,
            double height,
            double angle,
            string lineColor,
            string fillColor,
            int linePattern,
            int fillPattern,
            double lineWeight,
            bool isContainer,
            bool isCallout,
            bool isBackgroundSurface,
            bool isDiagramAdornment,
            string? calloutTargetId,
            IReadOnlyList<string> layers,
            IReadOnlyList<VisioInspectionShapeDataSnapshot> shapeData,
            IReadOnlyList<VisioInspectionUserCellSnapshot> userCells,
            IReadOnlyDictionary<string, string> data,
            IReadOnlyList<VisioInspectionConnectionPointSnapshot> connectionPoints,
            IReadOnlyList<string> childIds) {
            Id = id;
            Name = name;
            NameU = nameU;
            Type = type;
            MasterId = masterId;
            MasterNameU = masterNameU;
            MasterShapeId = masterShapeId;
            ParentId = parentId;
            Text = text;
            PinX = pinX;
            PinY = pinY;
            Width = width;
            Height = height;
            Angle = angle;
            LineColor = lineColor;
            FillColor = fillColor;
            LinePattern = linePattern;
            FillPattern = fillPattern;
            LineWeight = lineWeight;
            IsContainer = isContainer;
            IsCallout = isCallout;
            IsBackgroundSurface = isBackgroundSurface;
            IsDiagramAdornment = isDiagramAdornment;
            CalloutTargetId = calloutTargetId;
            Layers = layers;
            ShapeData = shapeData;
            UserCells = userCells;
            Data = data;
            ConnectionPoints = connectionPoints;
            ChildIds = childIds;
        }

        /// <summary>Shape identifier.</summary>
        public string Id { get; }

        /// <summary>Shape display name.</summary>
        public string? Name { get; }

        /// <summary>Shape universal name.</summary>
        public string? NameU { get; }

        /// <summary>Visio shape type, such as Group, when available.</summary>
        public string? Type { get; }

        /// <summary>Referenced master identifier.</summary>
        public string? MasterId { get; }

        /// <summary>Referenced master universal name.</summary>
        public string? MasterNameU { get; }

        /// <summary>Referenced master shape identifier.</summary>
        public string? MasterShapeId { get; }

        /// <summary>Parent group shape identifier.</summary>
        public string? ParentId { get; }

        /// <summary>Shape text.</summary>
        public string? Text { get; }

        /// <summary>Shape pin X coordinate.</summary>
        public double PinX { get; }

        /// <summary>Shape pin Y coordinate.</summary>
        public double PinY { get; }

        /// <summary>Shape width.</summary>
        public double Width { get; }

        /// <summary>Shape height.</summary>
        public double Height { get; }

        /// <summary>Shape rotation angle in radians.</summary>
        public double Angle { get; }

        /// <summary>Line color as a stable OfficeIMO color string.</summary>
        public string LineColor { get; }

        /// <summary>Fill color as a stable OfficeIMO color string.</summary>
        public string FillColor { get; }

        /// <summary>Visio line pattern value.</summary>
        public int LinePattern { get; }

        /// <summary>Visio fill pattern value.</summary>
        public int FillPattern { get; }

        /// <summary>Shape line weight.</summary>
        public double LineWeight { get; }

        /// <summary>Whether the shape is marked as a Visio container.</summary>
        public bool IsContainer { get; }

        /// <summary>Whether the shape is marked as an OfficeIMO callout.</summary>
        public bool IsCallout { get; }

        /// <summary>Whether the shape is marked as a background surface.</summary>
        public bool IsBackgroundSurface { get; }

        /// <summary>Whether the shape is marked as generated diagram adornment.</summary>
        public bool IsDiagramAdornment { get; }

        /// <summary>Callout target shape identifier, when present.</summary>
        public string? CalloutTargetId { get; }

        /// <summary>Layer names assigned to the shape.</summary>
        public IReadOnlyList<string> Layers { get; }

        /// <summary>Shape Data rows attached to the shape.</summary>
        public IReadOnlyList<VisioInspectionShapeDataSnapshot> ShapeData { get; }

        /// <summary>User cell rows attached to the shape.</summary>
        public IReadOnlyList<VisioInspectionUserCellSnapshot> UserCells { get; }

        /// <summary>Arbitrary data attached to the shape.</summary>
        public IReadOnlyDictionary<string, string> Data { get; }

        /// <summary>Connection points attached to the shape.</summary>
        public IReadOnlyList<VisioInspectionConnectionPointSnapshot> ConnectionPoints { get; }

        /// <summary>Number of connection points attached to the shape.</summary>
        public int ConnectionPointCount => ConnectionPoints.Count;

        /// <summary>Child shape identifiers when this shape is a group.</summary>
        public IReadOnlyList<string> ChildIds { get; }

        internal void AppendText(StringBuilder builder, string pagePrefix) {
            string prefix = pagePrefix + ".shape[" + VisioInspectionSnapshot.EscapeKey(Id) + "]";
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".name", Name);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".nameU", NameU);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".type", Type);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".masterId", MasterId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".masterNameU", MasterNameU);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".masterShapeId", MasterShapeId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".parentId", ParentId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".text", Text);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".pinX", PinX);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".pinY", PinY);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".width", Width);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".height", Height);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".angle", Angle);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".lineColor", LineColor);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".fillColor", FillColor);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".linePattern", LinePattern);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".fillPattern", FillPattern);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".lineWeight", LineWeight);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".isContainer", IsContainer);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".isCallout", IsCallout);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".isBackgroundSurface", IsBackgroundSurface);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".isDiagramAdornment", IsDiagramAdornment);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".calloutTargetId", CalloutTargetId);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".layers", string.Join(",", Layers));
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".connectionPointCount", ConnectionPointCount);
            AppendConnectionPoints(builder, prefix, ConnectionPoints);
            VisioInspectionSnapshot.AppendLine(builder, prefix + ".children", string.Join(",", ChildIds));
            AppendShapeData(builder, prefix, ShapeData);
            AppendUserCells(builder, prefix, UserCells);
            AppendData(builder, prefix, Data);
        }

        internal static void AppendConnectionPoints(StringBuilder builder, string prefix, IReadOnlyList<VisioInspectionConnectionPointSnapshot> points) {
            foreach (VisioInspectionConnectionPointSnapshot point in points) {
                string pointPrefix = prefix + ".connectionPoint[" + point.Index.ToString(CultureInfo.InvariantCulture) + "]";
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".sectionIndex", point.SectionIndex);
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".x", point.X);
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".y", point.Y);
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".dirX", point.DirX);
                VisioInspectionSnapshot.AppendLine(builder, pointPrefix + ".dirY", point.DirY);
            }
        }

        internal static void AppendShapeData(StringBuilder builder, string prefix, IReadOnlyList<VisioInspectionShapeDataSnapshot> rows) {
            foreach (VisioInspectionShapeDataSnapshot row in rows) {
                string rowPrefix = prefix + ".shapeData[" + VisioInspectionSnapshot.EscapeKey(row.Name) + "]";
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".label", row.Label);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".value", row.Value);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".type", row.Type);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".format", row.Format);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".prompt", row.Prompt);
            }
        }

        internal static void AppendUserCells(StringBuilder builder, string prefix, IReadOnlyList<VisioInspectionUserCellSnapshot> rows) {
            foreach (VisioInspectionUserCellSnapshot row in rows) {
                string rowPrefix = prefix + ".user[" + VisioInspectionSnapshot.EscapeKey(row.Name) + "]";
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".value", row.Value);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".formula", row.Formula);
                VisioInspectionSnapshot.AppendLine(builder, rowPrefix + ".prompt", row.Prompt);
            }
        }

        internal static void AppendData(StringBuilder builder, string prefix, IReadOnlyDictionary<string, string> data) {
            foreach (KeyValuePair<string, string> pair in data.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)) {
                VisioInspectionSnapshot.AppendLine(builder, prefix + ".data[" + VisioInspectionSnapshot.EscapeKey(pair.Key) + "]", pair.Value);
            }
        }
    }
}
