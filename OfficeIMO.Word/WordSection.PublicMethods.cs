using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordSection {
        public WordSection SetMargins(PageMargin pageMargins) {
            var pageMargin = _sectionProperties.GetFirstChild<PageMargin>();
            if (pageMargin == null) {
                _sectionProperties.Append(pageMargins);
                // pageMargin = _sectionProperties.GetFirstChild<PageMargin>();
            } else {
                pageMargin.Remove();
                _sectionProperties.Append(pageMargins);
            }

            return this;
        }

        public WordParagraph AddParagraph(string text = "") {
            //WordParagraph wordParagraph = new WordParagraph(_document, true);
            //rdParagraph.Text = text;

            //if (this._paragraph == null) {
            //    this._document.AddParagraph(wordParagraph);
            //} else {

            //    // this._paragraph.InsertBeforeSelf(wordParagraph._paragraph);
            //    WordParagraph lastParagraphWithinSection = this.Paragraphs.Last();
            //    wordParagraph = lastParagraphWithinSection.AddParagraphAfterSelf();
            //    wordParagraph.Text = text;
            //    //paragraph._section = this;

            //}
            //return wordParagraph;


            //if (this.Paragraphs.Count == 0) {
            //    //WordParagraph paragraph = this._document.AddParagraph(text);
            //    WordParagraph paragraph = new WordParagraph(_document, true);
            //    paragraph.Text = text;

            //    this._paragraph.InsertBeforeSelf(paragraph._paragraph);

            //    //paragraph._section = this;
            //    return paragraph;
            //} else {
            //    WordParagraph lastParagraphWithinSection = this.Paragraphs.Last();

            //    WordParagraph paragraph = lastParagraphWithinSection.AddParagraphAfterSelf(this);
            //    paragraph._document = this._document;
            //    // paragraph._section = this;
            //    //this.Paragraphs.Add(paragraph);
            //    paragraph.Text = text;
            //    return paragraph;
            //}

            if (this.Paragraphs.Count == 0) {
                WordParagraph paragraph = this._document.AddParagraph(text);
                //paragraph._section = this;
                return paragraph;
            } else {
                WordParagraph lastParagraphWithinSection = this.Paragraphs.Last();

                WordParagraph paragraph = lastParagraphWithinSection.AddParagraphAfterSelf(this);
                paragraph._document = this._document;
                // paragraph._section = this;
                //this.Paragraphs.Add(paragraph);
                paragraph.Text = text;
                return paragraph;
            }
        }

        public WordWatermark AddWatermark(WordWatermarkStyle watermarkStyle, string text) {
            return new WordWatermark(this._document, this, this.Header.Default, watermarkStyle, text);
        }

        public WordSection SetBorders(WordBorder wordBorder) {
            this.Borders.SetBorder(wordBorder);

            return this;
        }

        public WordParagraph AddHorizontalLine(BorderValues lineType = BorderValues.Single, System.Drawing.Color? color = null, uint size = 12, uint space = 1) {
            return this.AddParagraph().AddHorizontalLine(lineType, color, size, space);
        }
    }
}