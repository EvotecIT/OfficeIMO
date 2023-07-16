using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {

    /**
    * postion of a character in a paragrapho
   * 1st ParagraphPositon
   * 2nd TextPosition
   * 3rd CharacterPosition 
   */
    internal class WordPositionInParagraph {
        private int posParagraph = 0, posText = 0, posChar = 0;

        public WordPositionInParagraph() {
        }

        public WordPositionInParagraph(int posRun, int posText, int posChar) {
            this.posParagraph = posRun;
            this.posChar = posChar;
            this.posText = posText;
        }

        public int Paragraph {
            get {
                return posParagraph;
            }
            set {
                this.posParagraph = value;
            }
        }

        public int Text {
            get {
                return posText;
            }
            set {
                this.posText = value;
            }
        }


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
