using DocumentFormat.OpenXml.Wordprocessing;
using Tabs = DocumentFormat.OpenXml.Wordprocessing.Tabs;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a tab stop within a <see cref="WordParagraph"/>.
    /// </summary>
    public class WordTabStop : System.IEquatable<WordTabStop> {

        private WordParagraph _paragraph { get; set; }

        private Tabs _tabs {
            get {
                if (_paragraph._paragraphProperties.Tabs == null) {
                    _paragraph._paragraphProperties.Append(new Tabs());
                }
                return _paragraph._paragraphProperties.Tabs;
            }
        }

        private TabStop _tabStop { get; set; }

        /// <summary>
        /// Gets or sets the alignment type for the tab stop.
        /// </summary>
        public TabStopValues Alignment {
            get {
                return _tabStop.Val;
            }
            set {
                _tabStop.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the leader character displayed before the tab stop.
        /// </summary>
        public TabStopLeaderCharValues Leader {
            get {
                return _tabStop.Leader;
            }
            set {
                _tabStop.Leader = value;
            }
        }

        /// <summary>
        /// Gets or sets the position of the tab stop in twentieths of a point.
        /// </summary>
        public int Position {
            get {
                return (int)_tabStop.Position;
            }
            set {
                _tabStop.Position = value;
            }
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="WordTabStop"/> class for the specified paragraph.
        /// </summary>
        /// <param name="wordParagraph">The paragraph to which this tab stop belongs.</param>
        public WordTabStop(WordParagraph wordParagraph) {
            _paragraph = wordParagraph;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTabStop"/> class using an existing <see cref="TabStop"/> element.
        /// </summary>
        /// <param name="wordParagraph">The paragraph to which this tab stop belongs.</param>
        /// <param name="tab">The underlying Open XML tab stop element.</param>
        public WordTabStop(WordParagraph wordParagraph, TabStop tab) {
            _paragraph = wordParagraph;
            _tabStop = tab;
        }

        /// <summary>
        /// Adds a new tab stop to the paragraph and sets this instance to reference it.
        /// </summary>
        /// <param name="position">The position of the tab stop in twentieths of a point.</param>
        /// <param name="alignment">Optional alignment value. Defaults to <see cref="TabStopValues.Left"/>.</param>
        /// <param name="leader">Optional leader value. Defaults to <see cref="TabStopLeaderCharValues.None"/>.</param>
        /// <returns>The current <see cref="WordTabStop"/> instance.</returns>
        internal WordTabStop AddTab(int position, TabStopValues? alignment = null, TabStopLeaderCharValues? leader = null) {
            alignment ??= TabStopValues.Left;
            leader ??= TabStopLeaderCharValues.None;
            TabStop tabStop1 = new TabStop() { Val = alignment, Leader = leader, Position = position };
            _tabs.Append(tabStop1);
            _tabStop = tabStop1;
            return this;
        }

        /// <summary>
        /// Determines whether the specified <see cref="WordTabStop"/> is equal to the current instance.
        /// </summary>
        /// <param name="other">The tab stop to compare with the current instance.</param>
        /// <returns><c>true</c> if the tab stops have the same alignment, leader and position; otherwise, <c>false</c>.</returns>
        public bool Equals(WordTabStop other) {
            if (other is null) return false;
            return Alignment == other.Alignment && Leader == other.Leader && Position == other.Position;
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current instance.
        /// </summary>
        /// <param name="obj">The object to compare with the current instance.</param>
        /// <returns><c>true</c> if the objects are equal; otherwise, <c>false</c>.</returns>
        public override bool Equals(object obj) {
            return obj is WordTabStop other && Equals(other);
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>A hash code for the current object.</returns>
        public override int GetHashCode() {
            unchecked {
                int hash = 17;
                hash = hash * 31 + Alignment.GetHashCode();
                hash = hash * 31 + Leader.GetHashCode();
                hash = hash * 31 + Position.GetHashCode();
                return hash;
            }
        }
    }
}
