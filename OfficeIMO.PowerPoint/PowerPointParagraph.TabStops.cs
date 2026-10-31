using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointParagraph {
        internal const int MaximumTabStops = 32;

        /// <summary>Gets the paragraph's explicit tab stops as physical lengths.</summary>
        public IReadOnlyList<PowerPointTabStop> TabStops => Paragraph.ParagraphProperties?
            .GetFirstChild<A.TabStopList>()?.Elements<A.TabStop>()
            .Select(tab => new PowerPointTabStop(global::OfficeIMO.OfficeLength.Emus(tab.Position?.Value ?? 0),
                FromTabAlignment(tab.Alignment?.Value ?? A.TextTabAlignmentValues.Left))).ToArray()
            ?? Array.Empty<PowerPointTabStop>();

        /// <summary>Replaces up to 32 explicit tab stops. An empty sequence removes locally configured stops.</summary>
        public PowerPointParagraph SetTabStops(IEnumerable<PowerPointTabStop> tabStops) {
            if (tabStops == null) throw new ArgumentNullException(nameof(tabStops));
            var list = new A.TabStopList();
            foreach (PowerPointTabStop tab in tabStops) {
                if (list.ChildElements.Count == MaximumTabStops)
                    throw new ArgumentException("DrawingML paragraphs support at most 32 explicit tab stops.", nameof(tabStops));
                if (tab == null) throw new ArgumentException("Tab stops cannot contain null.", nameof(tabStops));
                list.Append(new A.TabStop {
                    Position = checked((int)tab.Position.ToEmus()),
                    Alignment = tab.Alignment switch {
                        PowerPointTabAlignment.Center => A.TextTabAlignmentValues.Center,
                        PowerPointTabAlignment.Right => A.TextTabAlignmentValues.Right,
                        PowerPointTabAlignment.Decimal => A.TextTabAlignmentValues.Decimal,
                        _ => A.TextTabAlignmentValues.Left
                    }
                });
            }
            var properties = EnsureParagraphProperties();
            properties.RemoveAllChildren<A.TabStopList>();
            if (list.HasChildren) InsertParagraphPropertyChild(properties, list);
            return this;
        }

        private static PowerPointTabAlignment FromTabAlignment(A.TextTabAlignmentValues alignment) {
            if (alignment == A.TextTabAlignmentValues.Left) return PowerPointTabAlignment.Left;
            if (alignment == A.TextTabAlignmentValues.Center) return PowerPointTabAlignment.Center;
            if (alignment == A.TextTabAlignmentValues.Right) return PowerPointTabAlignment.Right;
            if (alignment == A.TextTabAlignmentValues.Decimal) return PowerPointTabAlignment.Decimal;
            throw new NotSupportedException("The paragraph contains an unsupported tab alignment.");
        }
    }

    /// <summary>Alignment of text at a custom paragraph tab stop.</summary>
    public enum PowerPointTabAlignment {
        /// <summary>Aligns the left edge.</summary>
        Left,
        /// <summary>Centers text.</summary>
        Center,
        /// <summary>Aligns the right edge.</summary>
        Right,
        /// <summary>Aligns the decimal separator.</summary>
        Decimal
    }

    /// <summary>An immutable explicit paragraph tab stop.</summary>
    public sealed class PowerPointTabStop {
        /// <summary>Creates a tab stop at a nonnegative position measured as physical lengths.</summary>
        public PowerPointTabStop(global::OfficeIMO.OfficeLength position, PowerPointTabAlignment alignment = PowerPointTabAlignment.Left) {
            if (position.ToEmus() < 0 || position.ToEmus() > int.MaxValue)
                throw new ArgumentOutOfRangeException(nameof(position));
            if (!Enum.IsDefined(typeof(PowerPointTabAlignment), alignment))
                throw new ArgumentOutOfRangeException(nameof(alignment));
            Position = position;
            Alignment = alignment;
        }
        /// <summary>Gets the position as physical lengths.</summary>
        public global::OfficeIMO.OfficeLength Position { get; }
        /// <summary>Gets the text alignment.</summary>
        public PowerPointTabAlignment Alignment { get; }
    }
}
