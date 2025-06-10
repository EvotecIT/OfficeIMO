using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {
    internal class WordTextSegment {
        private WordPositionInParagraph beginPos;
        private WordPositionInParagraph endPos;

        public WordTextSegment() {
            this.beginPos = new WordPositionInParagraph();
            this.endPos = new WordPositionInParagraph();
        }

        public WordTextSegment(int beginRun, int endRun, int beginText, int endText, int beginChar, int endChar) {
            WordPositionInParagraph beginPos = new WordPositionInParagraph(beginRun, beginText, beginChar);
            WordPositionInParagraph endPos = new WordPositionInParagraph(endRun, endText, endChar);
            this.beginPos = beginPos;
            this.endPos = endPos;
        }

        public WordTextSegment(WordPositionInParagraph beginPos, WordPositionInParagraph endPos) {
            this.beginPos = beginPos;
            this.endPos = endPos;
        }

        public WordPositionInParagraph BeginPos {
            get {
                return beginPos;
            }
            set {
                beginPos = value;
            }
        }

        public WordPositionInParagraph EndPos {
            get {
                return endPos;
            }
        }
        /// <summary>
        /// The index of the begin paragraph
        /// </summary>
        public int BeginIndex {
            get {
                return beginPos.Paragraph;
            }
            set {
                beginPos.Paragraph = value;
            }
        }
      
        /// <summary>
        /// The index of the start text character
        /// </summary>
        public int BeginChar {
            get {
                return beginPos.Char;
            }
            set {
                beginPos.Char = value;
            }
        }

        /// <summary>
        /// The index of the end paragraph
        /// </summary>
        public int EndIndex {
            get {
                return endPos.Paragraph;
            }
            set {
                endPos.Paragraph = value;
            }
        }
       
        /// <summary>
        /// the index of the end text character
        /// </summary>
        public int EndChar {
            get {
                return endPos.Char;
            }
            set {
                endPos.Char = value;
            }
        }
    }
}
