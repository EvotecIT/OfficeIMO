using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    /// <summary>
    /// Encapsulates a mathematical equation contained in a paragraph.
    /// Provides access to the underlying Open XML elements so the
    /// equation can be removed from the document when required.
    /// </summary>
    public class WordEquation : WordElement {
        private WordDocument _document;
        private Paragraph _paragraph;
        //private DocumentFormat.OpenXml.Math.Paragraph mathParagraph;
        private readonly DocumentFormat.OpenXml.Math.OfficeMath? _officeMath;
        private readonly DocumentFormat.OpenXml.Math.Paragraph? _mathParagraph;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordEquation"/> class using
        /// the specified <paramref name="document"/>, <paramref name="paragraph"/>,
        /// and <paramref name="officeMath"/> equation element.
        /// </summary>
        /// <param name="document">Parent Word document.</param>
        /// <param name="paragraph">Paragraph that contains the equation.</param>
        /// <param name="officeMath">Math equation element.</param>
        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
            this._document = document;
            this._paragraph = paragraph;
            this._officeMath = officeMath;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordEquation"/> class using
        /// the specified <paramref name="document"/>, <paramref name="paragraph"/>,
        /// and <paramref name="mathParagraph"/> paragraph element.
        /// </summary>
        /// <param name="document">Parent Word document.</param>
        /// <param name="paragraph">Paragraph that contains the equation.</param>
        /// <param name="mathParagraph">Paragraph element representing the equation.</param>
        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            this._document = document;
            this._paragraph = paragraph;
            this._mathParagraph = mathParagraph;

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordEquation"/> class using
        /// both math equation elements: <paramref name="officeMath"/> and
        /// <paramref name="mathParagraph"/>.
        /// </summary>
        /// <param name="document">Parent Word document.</param>
        /// <param name="paragraph">Paragraph that contains the equation.</param>
        /// <param name="officeMath">Math equation element.</param>
        /// <param name="mathParagraph">Paragraph element representing the equation.</param>
        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            this._document = document;
            this._paragraph = paragraph;
            this._officeMath = officeMath;
            this._mathParagraph = mathParagraph;
        }

        /// <summary>
        /// Removes the equation from the document by deleting the underlying
        /// Open XML elements.
        /// </summary>
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
