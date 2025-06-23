namespace OfficeIMO.Word {
    public partial class WordParagraph : System.IEquatable<WordParagraph> {
        public bool Equals(WordParagraph other) {
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

        public override bool Equals(object obj) {
            return obj is WordParagraph other && Equals(other);
        }

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

        public static bool operator ==(WordParagraph left, WordParagraph right) {
            if (left is null) return right is null;
            return left.Equals(right);
        }

        public static bool operator !=(WordParagraph left, WordParagraph right) {
            return !(left == right);
        }
    }
}
