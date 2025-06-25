using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a continuous text fragment within a document and exposes
    /// its start and end positions.
    /// </summary>
    internal class WordTextSegment {
        private WordPositionInParagraph beginPos;
        private WordPositionInParagraph endPos;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTextSegment"/> class
        /// with both positions set to the start of the document.
        /// </summary>
        public WordTextSegment() {
            this.beginPos = new WordPositionInParagraph();
            this.endPos = new WordPositionInParagraph();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTextSegment"/> class
        /// with explicit position values.
        /// </summary>
        /// <param name="beginRun">Paragraph index for the start of the segment.</param>
        /// <param name="endRun">Paragraph index where the segment ends.</param>
        /// <param name="beginText">Text index within the starting paragraph.</param>
        /// <param name="endText">Text index within the ending paragraph.</param>
        /// <param name="beginChar">Character index within the starting text.</param>
        /// <param name="endChar">Character index within the ending text.</param>
        public WordTextSegment(int beginRun, int endRun, int beginText, int endText, int beginChar, int endChar) {
            WordPositionInParagraph beginPos = new WordPositionInParagraph(beginRun, beginText, beginChar);
            WordPositionInParagraph endPos = new WordPositionInParagraph(endRun, endText, endChar);
            this.beginPos = beginPos;
            this.endPos = endPos;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTextSegment"/> class
        /// using the supplied begin and end positions.
        /// </summary>
        /// <param name="beginPos">Starting position of the segment.</param>
        /// <param name="endPos">Ending position of the segment.</param>
        public WordTextSegment(WordPositionInParagraph beginPos, WordPositionInParagraph endPos) {
            this.beginPos = beginPos;
            this.endPos = endPos;
        }

        /// <summary>
        /// Gets or sets the starting position of the text segment.
        /// </summary>
        public WordPositionInParagraph BeginPos {
            get {
                return beginPos;
            }
            set {
                beginPos = value;
            }
        }

        /// <summary>
        /// Gets the ending position of the text segment.
        /// </summary>
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
