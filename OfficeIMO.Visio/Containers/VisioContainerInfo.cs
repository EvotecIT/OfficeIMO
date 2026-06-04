using System.Collections.Generic;
using System.Collections.ObjectModel;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Snapshot of typed Visio-native container metadata and current membership.
    /// </summary>
    public sealed class VisioContainerInfo {
        internal VisioContainerInfo(
            string id,
            string? text,
            IReadOnlyList<string> memberIds,
            double margin,
            double headingHeight,
            bool autoResize,
            bool locked,
            bool noHighlight,
            bool noRibbon,
            int containerStyle,
            int headingStyle,
            Color fillColor,
            Color lineColor,
            double lineWeight,
            int fillPattern,
            int linePattern,
            VisioTextStyle? textStyle) {
            Id = id;
            Text = text;
            MemberIds = new ReadOnlyCollection<string>(new List<string>(memberIds));
            Margin = margin;
            HeadingHeight = headingHeight;
            AutoResize = autoResize;
            Locked = locked;
            NoHighlight = noHighlight;
            NoRibbon = noRibbon;
            ContainerStyle = containerStyle;
            HeadingStyle = headingStyle;
            FillColor = fillColor;
            LineColor = lineColor;
            LineWeight = lineWeight;
            FillPattern = fillPattern;
            LinePattern = linePattern;
            TextStyle = textStyle?.Clone();
        }

        /// <summary>Container shape identifier.</summary>
        public string Id { get; }

        /// <summary>Container heading text.</summary>
        public string? Text { get; }

        /// <summary>Shape identifiers currently referenced by the container.</summary>
        public IReadOnlyList<string> MemberIds { get; }

        /// <summary>Number of current member shape identifiers.</summary>
        public int MemberCount => MemberIds.Count;

        /// <summary>Outer member margin in the owning page unit.</summary>
        public double Margin { get; }

        /// <summary>Heading height in the owning page unit.</summary>
        public double HeadingHeight { get; }

        /// <summary>Whether Visio may resize the container around members.</summary>
        public bool AutoResize { get; }

        /// <summary>Whether the container is marked locked in native container metadata.</summary>
        public bool Locked { get; }

        /// <summary>Whether selection highlighting is suppressed for the container.</summary>
        public bool NoHighlight { get; }

        /// <summary>Whether container ribbon UI is suppressed for the container.</summary>
        public bool NoRibbon { get; }

        /// <summary>Native Visio container style identifier.</summary>
        public int ContainerStyle { get; }

        /// <summary>Native Visio heading style identifier.</summary>
        public int HeadingStyle { get; }

        /// <summary>Container fill color.</summary>
        public Color FillColor { get; }

        /// <summary>Container border color.</summary>
        public Color LineColor { get; }

        /// <summary>Container border weight in inches.</summary>
        public double LineWeight { get; }

        /// <summary>Container fill pattern.</summary>
        public int FillPattern { get; }

        /// <summary>Container line pattern.</summary>
        public int LinePattern { get; }

        /// <summary>Container heading text style, when modeled.</summary>
        public VisioTextStyle? TextStyle { get; }
    }
}
