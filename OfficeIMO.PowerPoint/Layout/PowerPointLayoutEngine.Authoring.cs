using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointLayoutEngine {
        /// <summary>Uses the same measured content zone and column allocation as the semantic renderer.</summary>
        internal static PowerPointLayoutRect[] AuthoringFrames(PowerPointLayoutContext context,
            PowerPointLayoutSlide content, int count, double titleBottom) {
            if (count < 1 || count > 2) throw new ArgumentOutOfRangeException(nameof(count));
            PowerPointLayoutRect zone = ContentZone(context, MeasureHeader(context, content));
            if (titleBottom > zone.Y - 12D) zone = zone.Inset(0, titleBottom + 12D - zone.Y, 0, 0);
            if (zone.Width <= 0D || zone.Height <= 0D)
                throw new InvalidOperationException("The title leaves no content frame. Use another slide or explicit geometry.");
            if (count == 1) return new[] { zone };
            return TwoContentFrames(context, zone, false, false);
        }

        /// <summary>Shared allocation for semantic TwoContent and attached ordinary authoring.</summary>
        private static PowerPointLayoutRect[] TwoContentFrames(PowerPointLayoutContext context,
            PowerPointLayoutRect zone, bool leftText, bool rightText) {
            (int Column, int Span)[] columns;
            if (leftText != rightText) {
                columns = context.Grid.IsNarrow
                    ? leftText ? new[] { (1, 5), (6, 7) } : new[] { (1, 7), (8, 5) }
                    : leftText ? new[] { (1, 4), (6, 7) } : new[] { (1, 7), (9, 4) };
            } else {
                columns = new[] { (1, 6), (7, 6) };
            }
            return columns.Select(column => new PowerPointLayoutRect(context.Grid.X(column.Column), zone.Y,
                context.Grid.W(column.Span), zone.Height)).ToArray();
        }

        internal static PowerPointLayoutBox PhysicalBox(PowerPointLayoutRect box) => new(
            OfficeLength.Points(box.X), OfficeLength.Points(box.Y), OfficeLength.Points(box.Width), OfficeLength.Points(box.Height));
    }
}
