using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a character position inside a paragraph.
    /// </summary>
    internal class WordPositionInParagraph {
        private int posParagraph = 0, posText = 0, posChar = 0;

        /// <summary>
        /// Initializes a new instance with all positions set to zero.
        /// </summary>
        public WordPositionInParagraph() {
        }

        /// <summary>
        /// Initializes a new instance with the specified positions.
        /// </summary>
        /// <param name="posRun">Paragraph index.</param>
        /// <param name="posText">Text index.</param>
        /// <param name="posChar">Character index.</param>
        public WordPositionInParagraph(int posRun, int posText, int posChar) {
            this.posParagraph = posRun;
            this.posChar = posChar;
            this.posText = posText;
        }

        /// <summary>
        /// The paragraph index.
        /// </summary>
        public int Paragraph {
            get {
                return posParagraph;
            }
            set {
                this.posParagraph = value;
            }
        }

        /// <summary>
        /// The text index.
        /// </summary>
        public int Text {
            get {
                return posText;
            }
            set {
                this.posText = value;
            }
        }


        /// <summary>
        /// The character index.
        /// </summary>
        public int Char {
            get {
                return posChar;
            }
            set {
                this.posChar = value;
            }
        }
    }
}
