using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Reflection;
using System.Linq;
using System.Xml.Linq;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word {
    /// <summary>
    /// Contains public methods for editing paragraphs.
    /// </summary>
    public partial class WordParagraph {
        // Should the return type be changed to signify the difference between paragraph and run?
        /// <summary>
        /// Add a text to existing paragraph
        /// </summary>
        /// <param name="text">The text to be added to the paragraph.</param>
        /// <returns>The paragraph containing the new text.</returns>
        public WordParagraph AddText(string text) {
            WordParagraph wordParagraph = ConvertToTextWithBreaks(text);
            //WordParagraph wordParagraph = new WordParagraph(this._document, this._paragraph, new Run());
            //wordParagraph.Text = text;
            //this._paragraph.Append(wordParagraph._run);
            return wordParagraph;
        }

        /// <summary>
        /// Adds formatted text to the paragraph and applies basic run properties.
        /// </summary>
        /// <param name="text">Text to insert.</param>
        /// <param name="bold">Whether the text should be bold.</param>
        /// <param name="italic">Whether the text should be italic.</param>
        /// <param name="underline">Optional underline style.</param>
        /// <returns>The run containing the formatted text.</returns>
        public WordParagraph AddFormattedText(string text, bool bold = false, bool italic = false, UnderlineValues? underline = null) {
            var run = AddText(text);
            if (bold) {
                run.SetBold();
            }
            if (italic) {
                run.SetItalic();
            }
            if (underline != null) {
                run.SetUnderline(underline.Value);
            }
            return run;
        }
    }
}