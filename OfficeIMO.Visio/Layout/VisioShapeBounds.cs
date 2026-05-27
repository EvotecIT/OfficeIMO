using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Bounds of one or more Visio shapes in page coordinate units, expressed in inches.
    /// </summary>
    public readonly struct VisioShapeBounds : IEquatable<VisioShapeBounds> {
        /// <summary>
        /// Initializes a new bounds value.
        /// </summary>
        /// <param name="left">Left edge.</param>
        /// <param name="bottom">Bottom edge.</param>
        /// <param name="right">Right edge.</param>
        /// <param name="top">Top edge.</param>
        public VisioShapeBounds(double left, double bottom, double right, double top) {
            Left = left;
            Bottom = bottom;
            Right = right;
            Top = top;
            IsEmpty = false;
        }

        private VisioShapeBounds(bool isEmpty) {
            Left = 0;
            Bottom = 0;
            Right = 0;
            Top = 0;
            IsEmpty = isEmpty;
        }

        /// <summary>Empty bounds.</summary>
        public static VisioShapeBounds Empty => new VisioShapeBounds(true);

        /// <summary>Left edge.</summary>
        public double Left { get; }

        /// <summary>Bottom edge.</summary>
        public double Bottom { get; }

        /// <summary>Right edge.</summary>
        public double Right { get; }

        /// <summary>Top edge.</summary>
        public double Top { get; }

        /// <summary>Width of the bounds.</summary>
        public double Width => IsEmpty ? 0 : Right - Left;

        /// <summary>Height of the bounds.</summary>
        public double Height => IsEmpty ? 0 : Top - Bottom;

        /// <summary>Horizontal center of the bounds.</summary>
        public double CenterX => IsEmpty ? 0 : Left + Width / 2;

        /// <summary>Vertical center of the bounds.</summary>
        public double CenterY => IsEmpty ? 0 : Bottom + Height / 2;

        /// <summary>Whether the bounds contain no shapes.</summary>
        public bool IsEmpty { get; }

        /// <inheritdoc />
        public bool Equals(VisioShapeBounds other) {
            return Left.Equals(other.Left)
                && Bottom.Equals(other.Bottom)
                && Right.Equals(other.Right)
                && Top.Equals(other.Top)
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
                hash = (hash * 31) + Left.GetHashCode();
                hash = (hash * 31) + Bottom.GetHashCode();
                hash = (hash * 31) + Right.GetHashCode();
                hash = (hash * 31) + Top.GetHashCode();
                hash = (hash * 31) + IsEmpty.GetHashCode();
                return hash;
            }
        }

        /// <inheritdoc />
        public override string ToString() {
            return IsEmpty
                ? "Empty"
                : string.Format(System.Globalization.CultureInfo.InvariantCulture, "Left={0:0.###}, Bottom={1:0.###}, Right={2:0.###}, Top={3:0.###}", Left, Bottom, Right, Top);
        }
    }
}
