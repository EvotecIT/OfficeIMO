namespace OfficeIMO.Word {
    /// <summary>
    /// Implements equality comparison for paragraphs.
    /// </summary>
    public partial class WordParagraph : System.IEquatable<WordParagraph> {
        /// <summary>
        /// Determines whether the specified <see cref="WordParagraph"/> is equal to the current instance.
        /// </summary>
        /// <param name="other">The paragraph to compare with the current instance.</param>
        /// <returns><c>true</c> if the paragraphs are equal; otherwise, <c>false</c>.</returns>
        public bool Equals(WordParagraph? other) {
            if (other is null) return false;
            if (ReferenceEquals(this, other)) return true;
            if (ReferenceEquals(_paragraph, other._paragraph)) return true;

            if (Text != other.Text) return false;
            if (TabStops.Count != other.TabStops.Count) return false;
            for (int i = 0; i < TabStops.Count; i++) {
                if (!TabStops[i].Equals(other.TabStops[i])) return false;
            }

            return true;
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current instance.
        /// </summary>
        /// <param name="obj">The object to compare with the current instance.</param>
        /// <returns><c>true</c> if the objects are equal; otherwise, <c>false</c>.</returns>
        public override bool Equals(object? obj) {
            return obj is WordParagraph other && Equals(other);
        }

        /// <summary>
        /// Serves as the default hash function.
        /// </summary>
        /// <returns>A hash code for the current object.</returns>
        public override int GetHashCode() {
            if (_paragraph != null) return _paragraph.GetHashCode();
            unchecked {
                int hash = 17;
                hash = hash * 31 + (Text != null ? Text.GetHashCode() : 0);
                foreach (var tab in TabStops) {
                    hash = hash * 31 + tab.GetHashCode();
                }
                return hash;
            }
        }

        /// <summary>
        /// Determines whether two <see cref="WordParagraph"/> instances are equal.
        /// </summary>
        /// <param name="left">The left-hand instance.</param>
        /// <param name="right">The right-hand instance.</param>
        /// <returns><c>true</c> if the instances are equal; otherwise, <c>false</c>.</returns>
        public static bool operator ==(WordParagraph? left, WordParagraph? right) {
            if (left is null) return right is null;
            return left.Equals(right);
        }

        /// <summary>
        /// Determines whether two <see cref="WordParagraph"/> instances are not equal.
        /// </summary>
        /// <param name="left">The left-hand instance.</param>
        /// <param name="right">The right-hand instance.</param>
        /// <returns><c>true</c> if the instances are not equal; otherwise, <c>false</c>.</returns>
        public static bool operator !=(WordParagraph? left, WordParagraph? right) {
            return !(left == right);
        }
    }
}
