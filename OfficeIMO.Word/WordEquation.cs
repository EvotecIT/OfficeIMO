using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    public class WordEquation : WordElement {
        private WordDocument _document;
        private Paragraph _paragraph;
        private DocumentFormat.OpenXml.Math.Paragraph mathParagraph;
        private readonly DocumentFormat.OpenXml.Math.OfficeMath _officeMath;
        private readonly DocumentFormat.OpenXml.Math.Paragraph _mathParagraph;

        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
            this._document = document;
            this._paragraph = paragraph;
            this._officeMath = officeMath;
        }

        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            this._document = document;
            this._paragraph = paragraph;
            this._mathParagraph = mathParagraph;

        }

        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            this._document = document;
            this._paragraph = paragraph;
            this._officeMath = officeMath;
            this._mathParagraph = mathParagraph;
        }

        public void Remove() {
            if (this._officeMath != null) {
                this._officeMath.Remove();
            }

            if (this._mathParagraph != null) {
                this._mathParagraph.Remove();
            }
        }
    }
}
