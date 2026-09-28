using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static double ReserveOutsideRadialLabelGutters(double radius, double visualWidth) {
        // Both sides need the maximum supported label width, a leader gap, and frame padding.
        double availableRadius = (visualWidth - 174D) / 2D;
        if (availableRadius <= 0D)
            throw new NotSupportedException("Outside radial chart labels need a wider plot area.");
        return Math.Min(radius, availableRadius);
    }

    private sealed class RadialOutsideLabel {
        internal RadialOutsideLabel(string text, double width, double height, double anchorX,
            double anchorY, bool rightSide, bool hasSlice) {
            Text = text;
            Width = width;
            Height = height;
            AnchorX = anchorX;
            AnchorY = anchorY;
            RightSide = rightSide;
            HasSlice = hasSlice;
        }

        internal string Text { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal double AnchorX { get; }
        internal double AnchorY { get; }
        internal bool RightSide { get; }
        internal bool HasSlice { get; }
        internal double X { get; set; }
        internal double Y { get; set; }
    }

    private static void QueueRadialOutsideLabel(List<RadialOutsideLabel> labels,
        OfficeChartLayout layout, string category, OfficeChartSeries series, double value,
        double total, double centerX, double centerY, double outerRadius, double angle,
        bool hasSlice, double percentageRatio) {
        string text = FormatDataLabel(layout, category, series, value, total, percentageRatio);
        if (string.IsNullOrWhiteSpace(text)) return;
        double labelWidth = Math.Min(78D, Math.Max(40D,
            text.Length * layout.DataLabelFontSize * 0.52D + 12D));
        double labelHeight = Math.Max(12D, layout.DataLabelFontSize + 6D);
        labels.Add(new RadialOutsideLabel(text, labelWidth, labelHeight,
            centerX + Math.Cos(angle) * outerRadius,
            centerY + Math.Sin(angle) * outerRadius,
            Math.Cos(angle) >= 0D, hasSlice));
    }

    private static void AddRadialOutsideLabels(OfficeDrawing drawing,
        IReadOnlyList<RadialOutsideLabel> labels, OfficeChartLayout layout, OfficeChartStyle style,
        double leftBound, double rightBound, double topBound, double bottomBound,
        double centerX, double edgeRadius) {
        if (labels.Count == 0) return;
        topBound = Math.Max(0D, topBound);
        bottomBound = Math.Min(drawing.Height, bottomBound);
        foreach (bool rightSide in new[] { false, true }) {
            RadialOutsideLabel[] side = labels.Where(label => label.RightSide == rightSide)
                .OrderBy(label => label.AnchorY).ToArray();
            if (side.Length == 0) continue;
            const double gap = 2D;
            double required = side.Sum(label => label.Height) + gap * (side.Length - 1);
            if (required > bottomBound - topBound)
                throw new NotSupportedException("Outside radial chart labels exceed the available height.");
            double cursor = topBound;
            foreach (RadialOutsideLabel label in side) {
                label.Y = Math.Max(label.AnchorY - label.Height / 2D, cursor);
                cursor = label.Y + label.Height + gap;
                double preferredX = rightSide
                    ? centerX + edgeRadius + 7D
                    : centerX - edgeRadius - 7D - label.Width;
                label.X = Math.Max(leftBound + 2D,
                    Math.Min(rightBound - label.Width - 2D, preferredX));
            }
            if (side[side.Length - 1].Y + side[side.Length - 1].Height > bottomBound) {
                side[side.Length - 1].Y = bottomBound - side[side.Length - 1].Height;
                for (int index = side.Length - 2; index >= 0; index--)
                    side[index].Y = Math.Min(side[index].Y,
                        side[index + 1].Y - side[index].Height - gap);
            }
            if (side[0].Y < topBound) {
                double shift = topBound - side[0].Y;
                foreach (RadialOutsideLabel label in side) label.Y += shift;
            }
        }
        OfficeColor lineColor = style.MutedTextColor;
        OfficeColor textColor = GetReadableDataLabelColor(
            style.DataLabelFillColor ?? style.PlotAreaBackgroundColor ?? style.BackgroundColor);
        foreach (RadialOutsideLabel label in labels) {
            if (!label.HasSlice || !layout.ShowDataLabelLeaderLines) continue;
            double targetX = label.RightSide ? label.X : label.X + label.Width;
            double targetY = label.Y + label.Height / 2D;
            AddPointLine(drawing, new[] {
                new OfficePoint(label.AnchorX, label.AnchorY),
                new OfficePoint(targetX, targetY)
            }, lineColor, 0.5D);
        }
        foreach (RadialOutsideLabel label in labels)
            AddDataLabel(drawing, layout, style, label.Text, label.X, label.Y,
                label.Width, label.Height, OfficeTextAlignment.Center, textColor);
    }
}
