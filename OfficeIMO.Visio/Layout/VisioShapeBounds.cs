using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Physical bounds of one or more Visio shapes in page coordinates.
    /// </summary>
    public readonly struct VisioShapeBounds : IEquatable<VisioShapeBounds> {
        /// <summary>
        /// Initializes a new bounds value.
        /// </summary>
        /// <param name="left">Left edge.</param>
        /// <param name="bottom">Bottom edge.</param>
        /// <param name="right">Right edge.</param>
        /// <param name="top">Top edge.</param>
        internal VisioShapeBounds(double left, double bottom, double right, double top) {
            LeftInches = left;
            BottomInches = bottom;
            RightInches = right;
            TopInches = top;
            IsEmpty = false;
        }

        /// <summary>Initializes bounds from physical page coordinates.</summary>
        public VisioShapeBounds(OfficeLength left, OfficeLength bottom, OfficeLength right, OfficeLength top)
            : this(left.ToInches(), bottom.ToInches(), right.ToInches(), top.ToInches()) { }

        private VisioShapeBounds(bool isEmpty) {
            LeftInches = 0;
            BottomInches = 0;
            RightInches = 0;
            TopInches = 0;
            IsEmpty = isEmpty;
        }

        /// <summary>Empty bounds.</summary>
        public static VisioShapeBounds Empty => new VisioShapeBounds(true);

        /// <summary>Left edge.</summary>
        internal double LeftInches { get; }

        /// <summary>Bottom edge.</summary>
        internal double BottomInches { get; }

        /// <summary>Right edge.</summary>
        internal double RightInches { get; }

        /// <summary>Top edge.</summary>
        internal double TopInches { get; }

        /// <summary>Width of the bounds.</summary>
        internal double WidthInches => IsEmpty ? 0 : RightInches - LeftInches;

        /// <summary>Height of the bounds.</summary>
        internal double HeightInches => IsEmpty ? 0 : TopInches - BottomInches;

        /// <summary>Horizontal center of the bounds.</summary>
        internal double CenterXInches => IsEmpty ? 0 : LeftInches + WidthInches / 2;

        /// <summary>Vertical center of the bounds.</summary>
        internal double CenterYInches => IsEmpty ? 0 : BottomInches + HeightInches / 2;

        /// <summary>Gets the physical left edge.</summary>
        public OfficeLength Left => OfficeLength.Inches(LeftInches);

        /// <summary>Gets the physical bottom edge.</summary>
        public OfficeLength Bottom => OfficeLength.Inches(BottomInches);

        /// <summary>Gets the physical right edge.</summary>
        public OfficeLength Right => OfficeLength.Inches(RightInches);

        /// <summary>Gets the physical top edge.</summary>
        public OfficeLength Top => OfficeLength.Inches(TopInches);

        /// <summary>Gets the physical width.</summary>
        public OfficeLength Width => OfficeLength.Inches(WidthInches);

        /// <summary>Gets the physical height.</summary>
        public OfficeLength Height => OfficeLength.Inches(HeightInches);

        /// <summary>Gets the horizontal center.</summary>
        public OfficeLength CenterX => OfficeLength.Inches(CenterXInches);

        /// <summary>Gets the vertical center.</summary>
        public OfficeLength CenterY => OfficeLength.Inches(CenterYInches);

        /// <summary>Whether the bounds contain no shapes.</summary>
        public bool IsEmpty { get; }

        /// <inheritdoc />
        public bool Equals(VisioShapeBounds other) {
            return LeftInches.Equals(other.LeftInches)
                && BottomInches.Equals(other.BottomInches)
                && RightInches.Equals(other.RightInches)
                && TopInches.Equals(other.TopInches)
                && IsEmpty == other.IsEmpty;
        }

        /// <inheritdoc />
        public override bool Equals(object? obj) {
            return obj is VisioShapeBounds other && Equals(other);
        }

        /// <inheritdoc />
        public override int GetHashCode() {
            unchecked {
                int hash = 17;
                hash = (hash * 31) + LeftInches.GetHashCode();
                hash = (hash * 31) + BottomInches.GetHashCode();
                hash = (hash * 31) + RightInches.GetHashCode();
                hash = (hash * 31) + TopInches.GetHashCode();
                hash = (hash * 31) + IsEmpty.GetHashCode();
                return hash;
            }
        }

        /// <inheritdoc />
        public override string ToString() {
            return IsEmpty
                ? "Empty"
                : string.Format(System.Globalization.CultureInfo.InvariantCulture, "Left={0:0.###}, Bottom={1:0.###}, Right={2:0.###}, Top={3:0.###}", LeftInches, BottomInches, RightInches, TopInches);
        }
    }
}
