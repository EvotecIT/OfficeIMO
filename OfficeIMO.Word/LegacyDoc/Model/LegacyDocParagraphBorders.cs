namespace OfficeIMO.Word.LegacyDoc.Model {
    internal enum LegacyDocParagraphBorderStyle {
        None,
        Single,
        Double,
        Dotted,
        Dashed
    }

    internal readonly struct LegacyDocParagraphBorder : IEquatable<LegacyDocParagraphBorder> {
        internal LegacyDocParagraphBorder(LegacyDocParagraphBorderStyle style, string? colorHex, int sizeEighthPoints, int spacePoints) {
            Style = style;
            ColorHex = string.IsNullOrWhiteSpace(colorHex)
                ? null
                : colorHex!.Replace("#", string.Empty).ToLowerInvariant();
            SizeEighthPoints = sizeEighthPoints;
            SpacePoints = spacePoints;
        }

        internal LegacyDocParagraphBorderStyle Style { get; }

        internal string? ColorHex { get; }

        internal int SizeEighthPoints { get; }

        internal int SpacePoints { get; }

        internal bool HasAny => Style != LegacyDocParagraphBorderStyle.None;

        public bool Equals(LegacyDocParagraphBorder other) {
            return Style == other.Style
                && string.Equals(ColorHex, other.ColorHex, StringComparison.Ordinal)
                && SizeEighthPoints == other.SizeEighthPoints
                && SpacePoints == other.SpacePoints;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocParagraphBorder other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Style.GetHashCode();
            hash = (hash * 31) + (ColorHex == null ? 0 : ColorHex.GetHashCode());
            hash = (hash * 31) + SizeEighthPoints.GetHashCode();
            hash = (hash * 31) + SpacePoints.GetHashCode();
            return hash;
        }
    }

    internal readonly struct LegacyDocParagraphBorders : IEquatable<LegacyDocParagraphBorders> {
        internal LegacyDocParagraphBorders(
            LegacyDocParagraphBorder top,
            LegacyDocParagraphBorder left,
            LegacyDocParagraphBorder bottom,
            LegacyDocParagraphBorder right,
            LegacyDocParagraphBorder between) {
            Top = top;
            Left = left;
            Bottom = bottom;
            Right = right;
            Between = between;
        }

        internal LegacyDocParagraphBorder Top { get; }

        internal LegacyDocParagraphBorder Left { get; }

        internal LegacyDocParagraphBorder Bottom { get; }

        internal LegacyDocParagraphBorder Right { get; }

        internal LegacyDocParagraphBorder Between { get; }

        internal bool HasAny => Top.HasAny || Left.HasAny || Bottom.HasAny || Right.HasAny || Between.HasAny;

        public bool Equals(LegacyDocParagraphBorders other) {
            return Top.Equals(other.Top)
                && Left.Equals(other.Left)
                && Bottom.Equals(other.Bottom)
                && Right.Equals(other.Right)
                && Between.Equals(other.Between);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocParagraphBorders other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Top.GetHashCode();
            hash = (hash * 31) + Left.GetHashCode();
            hash = (hash * 31) + Bottom.GetHashCode();
            hash = (hash * 31) + Right.GetHashCode();
            hash = (hash * 31) + Between.GetHashCode();
            return hash;
        }
    }
}
