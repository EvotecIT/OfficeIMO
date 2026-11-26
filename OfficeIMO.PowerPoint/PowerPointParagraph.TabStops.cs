using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointParagraph {
        internal const int MaximumTabStops = 32;

        /// <summary>Gets the paragraph's explicit tab stops as physical lengths.</summary>
        public IReadOnlyList<PowerPointTabStop> TabStops {
            get {
                PowerPointXmlNode list = Properties.Child(OfficeXmlNamespaces.DrawingML, "tabLst");
                if (list.IsNone) return Array.Empty<PowerPointTabStop>();
                var stops = new List<PowerPointTabStop>();
                foreach (PowerPointXmlNode tab in list.Children(OfficeXmlNamespaces.DrawingML, "tab")) {
                    stops.Add(new PowerPointTabStop(global::OfficeIMO.OfficeLength.Emus(tab.Int32Attribute("pos") ?? 0), FromTabAlignment(tab.Attribute("algn") ?? "l")));
                }
                return stops.ToArray();
            }
        }

        /// <summary>Replaces up to 32 explicit tab stops. An empty sequence removes locally configured stops.</summary>
        public PowerPointParagraph SetTabStops(IEnumerable<PowerPointTabStop> tabStops) {
            ArgumentNullException.ThrowIfNull(tabStops);
            var stops = new List<PowerPointTabStop>();
            foreach (PowerPointTabStop tab in tabStops) {
                if (stops.Count == MaximumTabStops)
                    throw new ArgumentException("DrawingML paragraphs support at most 32 explicit tab stops.", nameof(tabStops));
                if (tab == null) throw new ArgumentException("Tab stops cannot contain null.", nameof(tabStops));
                stops.Add(tab);
            }
            PowerPointXmlNode properties = EnsureProperties();
            properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "tabLst");
            if (stops.Count == 0) return this;
            PowerPointXmlNode list = properties.CreateElement(OfficeXmlNamespaces.DrawingML, "tabLst");
            foreach (PowerPointTabStop tab in stops) {
                PowerPointXmlNode stop = list.Append(list.CreateElement(OfficeXmlNamespaces.DrawingML, "tab"));
                stop.SetInt32Attribute("pos", checked((int)tab.Position.ToEmus()));
                stop.SetAttribute("algn", tab.Alignment switch {
                    PowerPointTabAlignment.Center => "ctr",
                    PowerPointTabAlignment.Right => "r",
                    PowerPointTabAlignment.Decimal => "dec",
                    _ => "l"
                });
            }
            properties.InsertOrdered(list, ATextParagraphPropertiesView.Order);
            return this;
        }

        private static PowerPointTabAlignment FromTabAlignment(string alignment) => alignment switch {
            "l" => PowerPointTabAlignment.Left,
            "ctr" => PowerPointTabAlignment.Center,
            "r" => PowerPointTabAlignment.Right,
            "dec" => PowerPointTabAlignment.Decimal,
            _ => throw new NotSupportedException("The paragraph contains an unsupported tab alignment.")
        };
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
