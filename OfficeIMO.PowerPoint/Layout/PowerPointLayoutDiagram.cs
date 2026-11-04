using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Draws an architecture diagram in a content zone: one lane per node group, evenly spaced boxes, and
    ///     connectors that run between box edges so they never cross a box or strike through their own label.
    /// </summary>
    internal static class PowerPointLayoutDiagram {
        internal static void Draw(PowerPointSlide slide, PowerPointLayoutContext context, PowerPointLayoutRect zone,
            PowerPointArchitectureContent content) {
            var writer = new PowerPointLayoutWriter(slide, context.Theme);
            PowerPointLayoutTheme theme = context.Theme;
            List<string?> groups = content.Nodes.Select(n => n.Group).Distinct().ToList();
            bool lanes = groups.Count > 1 || groups[0] != null;
            double labelWidth = lanes ? 120D : 0D;
            double laneGap = 16D;
            double laneHeight = Math.Min(130D, (zone.Height - (groups.Count - 1) * laneGap) / groups.Count);
            double top = zone.Y + Math.Max(0D, (zone.Height - groups.Count * laneHeight - (groups.Count - 1) * laneGap) / 2D);
            var boxes = new Dictionary<string, PowerPointLayoutRect>(StringComparer.OrdinalIgnoreCase);
            PowerPointTextStyleSpec heading = context.Style(PowerPointTextRole.Heading);
            PowerPointTextStyleSpec small = context.Style(PowerPointTextRole.BodySmall);
            for (int g = 0; g < groups.Count; g++) {
                var lane = new PowerPointLayoutRect(zone.X, top + g * (laneHeight + laneGap), zone.Width, laneHeight);
                if (lanes) {
                    writer.Rect("Architecture Lane", lane, theme.Surface);
                    writer.Text("Architecture Lane Label", new PowerPointLayoutRect(lane.X + 16D, lane.Y + 12D, labelWidth - 24D, lane.Height - 24D),
                        context.Paragraph(groups[g] ?? string.Empty, PowerPointTextRole.Eyebrow, theme.TextMuted));
                }

                List<PowerPointArchitectureNode> nodes = content.Nodes.Where(n => n.Group == groups[g]).ToList();
                double area = lane.Width - labelWidth - 24D;
                double gap = 48D;
                double width = Math.Min(220D, (area - (nodes.Count - 1) * gap) / Math.Max(1, nodes.Count));
                double used = nodes.Count * width + (nodes.Count - 1) * gap;
                double x = lane.X + labelWidth + (area - used) / 2D + 12D;
                double boxHeight = Math.Min(lane.Height - 24D, 84D);
                for (int i = 0; i < nodes.Count; i++) {
                    var box = new PowerPointLayoutRect(x + i * (width + gap), lane.Y + (lane.Height - boxHeight) / 2D, width, boxHeight);
                    boxes[nodes[i].Id] = box;
                    writer.Rect("Architecture Node", box, theme.Background, outline: theme.Hairline, outlineWidth: 1D, decorative: false);
                    writer.Rect("Architecture Node Rule", new PowerPointLayoutRect(box.X, box.Y, box.Width, 3D), theme.Accent1);
                    double titleSize = 16D;
                    while (titleSize > 12D && PowerPointTextMeasure.LongestWord(nodes[i].Title, titleSize, theme.HeadingFont, true) > box.Width - 24D) titleSize -= 1D;
                    var paragraphs = new List<PowerPointLayoutParagraph> { context.Paragraph(nodes[i].Title, PowerPointTextRole.Heading, size: titleSize) };
                    if (!string.IsNullOrWhiteSpace(nodes[i].Body)) {
                        PowerPointLayoutParagraph body = context.Paragraph(nodes[i].Body!, PowerPointTextRole.BodySmall, size: small.Minimum);
                        body.SpaceBefore = 4D;
                        paragraphs.Add(body);
                    }

                    writer.Text("Architecture Node Text", box.Inset(12D, 12D, 12D, 8D), paragraphs, anchor: "ctr");
                }
            }

            foreach (PowerPointArchitectureEdge edge in content.Edges) {
                if (!boxes.TryGetValue(edge.FromId, out PowerPointLayoutRect from) || !boxes.TryGetValue(edge.ToId, out PowerPointLayoutRect to)) continue;
                (double x1, double y1, double x2, double y2) = Endpoints(from, to);
                writer.Line("Architecture Connector", x1, y1, x2, y2, theme.TextMuted, 1.25D, arrow: true);
                if (string.IsNullOrWhiteSpace(edge.Label)) continue;
                PowerPointTextStyleSpec caption = context.Style(PowerPointTextRole.Caption);
                double labelW = PowerPointTextMeasure.Width(edge.Label!, caption.Size, theme.BodyFont) + 10D;
                double mx = (x1 + x2) / 2D, my = (y1 + y2) / 2D;
                var labelBox = new PowerPointLayoutRect(mx - labelW / 2D, my - caption.Size * 0.7D, labelW, caption.Size * 1.4D);
                // The label sits on a background-coloured plate so the connector never strikes through it.
                writer.Text("Architecture Connector Label", labelBox, new[] { context.Paragraph(edge.Label!, PowerPointTextRole.Caption, align: "ctr") },
                    anchor: "ctr", fill: theme.Background);
            }
        }

        /// <summary>Points on the facing edges of two boxes.</summary>
        private static (double, double, double, double) Endpoints(PowerPointLayoutRect a, PowerPointLayoutRect b) {
            double acx = a.X + a.Width / 2D, acy = a.Y + a.Height / 2D, bcx = b.X + b.Width / 2D, bcy = b.Y + b.Height / 2D;
            if (b.X >= a.Right) return (a.Right, acy, b.X, bcy);
            if (b.Right <= a.X) return (a.X, acy, b.Right, bcy);
            if (b.Y >= a.Bottom) return (acx, a.Bottom, bcx, b.Y);
            return (acx, a.Y, bcx, b.Bottom);
        }
    }
}
