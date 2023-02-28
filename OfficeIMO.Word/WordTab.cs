using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Wordprocessing;
using Tabs = DocumentFormat.OpenXml.Wordprocessing.Tabs;

namespace OfficeIMO.Word {
    public class WordTab {

        private WordParagraph Paragraph { get; set; }

        private Tabs Tabs {
            get {
                if (Paragraph._paragraphProperties.Tabs == null) {
                    Paragraph._paragraphProperties.Append(new Tabs());
                }
                return Paragraph._paragraphProperties.Tabs;
            }
        }
        private TabStop TabStop { get; set; }

        public TabStopValues Alignment {
            get {
                return TabStop.Val;
            }
            set {
                TabStop.Val = value;
            }
        }

        public TabStopLeaderCharValues Leader {
            get {
                return TabStop.Leader;
            }
            set {
                TabStop.Leader = value;
            }
        }

        public int Position {
            get {
                return (int)TabStop.Position;
            }
            set {
                TabStop.Position = value;
            }
        }


        public WordTab(WordParagraph wordParagraph) {
            Paragraph = wordParagraph;
        }

        public WordTab(WordParagraph wordParagraph, TabStop tab) {
            Paragraph = wordParagraph;
            TabStop = tab;
        }

        internal WordTab AddTab(int position, TabStopValues alignment = TabStopValues.Left, TabStopLeaderCharValues leader = TabStopLeaderCharValues.None) {
            TabStop tabStop1 = new TabStop() { Val = alignment, Leader = leader, Position = position };
            Tabs.Append(tabStop1);
            TabStop = tabStop1;
            return this;
        }
    }
}
