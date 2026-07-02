namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocTabStop : IEquatable<LegacyDocTabStop> {
        internal LegacyDocTabStop(int positionTwips, LegacyDocTabStopAlignment alignment, LegacyDocTabStopLeader leader) {
            PositionTwips = positionTwips;
            Alignment = alignment;
            Leader = leader;
        }

        internal int PositionTwips { get; }

        internal LegacyDocTabStopAlignment Alignment { get; }

        internal LegacyDocTabStopLeader Leader { get; }

        public bool Equals(LegacyDocTabStop other) {
            return PositionTwips == other.PositionTwips
                && Alignment == other.Alignment
                && Leader == other.Leader;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocTabStop other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + PositionTwips.GetHashCode();
            hash = (hash * 31) + Alignment.GetHashCode();
            hash = (hash * 31) + Leader.GetHashCode();
            return hash;
        }
    }

    internal enum LegacyDocTabStopAlignment {
        Left,
        Center,
        Right,
        Decimal,
        Bar,
        Clear
    }

    internal enum LegacyDocTabStopLeader {
        None,
        Dot,
        Hyphen,
        Underscore,
        Heavy,
        MiddleDot
    }
}
