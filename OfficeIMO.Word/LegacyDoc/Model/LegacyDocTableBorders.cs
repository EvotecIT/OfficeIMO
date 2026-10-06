namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Whole-table border defaults, distinct from explicit cell overrides.</summary>
    internal readonly struct LegacyDocTableBorders : IEquatable<LegacyDocTableBorders> {
        internal LegacyDocTableBorders(
            LegacyDocTableCellBorder top, LegacyDocTableCellBorder left,
            LegacyDocTableCellBorder bottom, LegacyDocTableCellBorder right,
            LegacyDocTableCellBorder insideHorizontal, LegacyDocTableCellBorder insideVertical) {
            Top = top;
            Left = left;
            Bottom = bottom;
            Right = right;
            InsideHorizontal = insideHorizontal;
            InsideVertical = insideVertical;
        }

        internal LegacyDocTableCellBorder Top { get; }
        internal LegacyDocTableCellBorder Left { get; }
        internal LegacyDocTableCellBorder Bottom { get; }
        internal LegacyDocTableCellBorder Right { get; }
        internal LegacyDocTableCellBorder InsideHorizontal { get; }
        internal LegacyDocTableCellBorder InsideVertical { get; }
        internal bool HasAny => Top.HasAny || Left.HasAny || Bottom.HasAny || Right.HasAny
            || InsideHorizontal.HasAny || InsideVertical.HasAny;

        /// <summary>Fills unspecified edges from the next lower formatting level.</summary>
        internal LegacyDocTableBorders WithDefaults(LegacyDocTableBorders defaults) => new LegacyDocTableBorders(
            Top.HasAny ? Top : defaults.Top, Left.HasAny ? Left : defaults.Left,
            Bottom.HasAny ? Bottom : defaults.Bottom, Right.HasAny ? Right : defaults.Right,
            InsideHorizontal.HasAny ? InsideHorizontal : defaults.InsideHorizontal,
            InsideVertical.HasAny ? InsideVertical : defaults.InsideVertical);

        public bool Equals(LegacyDocTableBorders other) => Top.Equals(other.Top) && Left.Equals(other.Left)
            && Bottom.Equals(other.Bottom) && Right.Equals(other.Right)
            && InsideHorizontal.Equals(other.InsideHorizontal) && InsideVertical.Equals(other.InsideVertical);
        public override bool Equals(object? obj) => obj is LegacyDocTableBorders other && Equals(other);
        public override int GetHashCode() {
            int hash = Top.GetHashCode();
            hash = hash * 31 + Left.GetHashCode();
            hash = hash * 31 + Bottom.GetHashCode();
            hash = hash * 31 + Right.GetHashCode();
            hash = hash * 31 + InsideHorizontal.GetHashCode();
            return hash * 31 + InsideVertical.GetHashCode();
        }
    }
}
