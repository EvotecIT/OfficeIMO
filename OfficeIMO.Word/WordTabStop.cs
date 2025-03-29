using DocumentFormat.OpenXml.Wordprocessing;
using Tabs = DocumentFormat.OpenXml.Wordprocessing.Tabs;

namespace OfficeIMO.Word {
    public class WordTabStop {

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

        public TabStopValues Alignment {
            get {
                return _tabStop.Val;
            }
            set {
                _tabStop.Val = value;
            }
        }

        public TabStopLeaderCharValues Leader {
            get {
                return _tabStop.Leader;
            }
            set {
                _tabStop.Leader = value;
            }
        }

        public int Position {
            get {
                return (int)_tabStop.Position;
            }
            set {
                _tabStop.Position = value;
            }
        }


        public WordTabStop(WordParagraph wordParagraph) {
            _paragraph = wordParagraph;
        }

        public WordTabStop(WordParagraph wordParagraph, TabStop tab) {
            _paragraph = wordParagraph;
            _tabStop = tab;
        }

        internal WordTabStop AddTab(int position, TabStopValues? alignment = null, TabStopLeaderCharValues? leader = null) {
            alignment ??= TabStopValues.Left;
            leader ??= TabStopLeaderCharValues.None;
            TabStop tabStop1 = new TabStop() { Val = alignment, Leader = leader, Position = position };
            _tabs.Append(tabStop1);
            _tabStop = tabStop1;
            return this;
        }
    }
}
