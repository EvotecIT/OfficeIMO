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
        public WordSection SetMargins(WordMargin pageMargins) {
            return WordMargins.SetMargins(this, pageMargins);
        }

        public WordParagraph AddParagraph(bool newRun) {
            var wordParagraph = new WordParagraph(_document, newParagraph: true, newRun: newRun);
            if (this.Paragraphs.Count == 0) {
                WordParagraph paragraph = this._document.AddParagraph(wordParagraph);
                return paragraph;
            } else {
                WordParagraph lastParagraphWithinSection = this.Paragraphs.Last();
                WordParagraph paragraph = lastParagraphWithinSection.AddParagraphAfterSelf(this, wordParagraph);
                return paragraph;
            }
        }

        public WordParagraph AddParagraph(string text = "") {
            if (this.Paragraphs.Count == 0) {
                WordParagraph paragraph = this._document.AddParagraph();
                if (text != "") {
                    paragraph.Text = text;
                }

                return paragraph;
            } else {
                WordParagraph lastParagraphWithinSection = this.Paragraphs.Last();

                WordParagraph paragraph = lastParagraphWithinSection.AddParagraphAfterSelf(this);
                paragraph._document = this._document;
                if (text != "") {
                    paragraph.Text = text;
                }

                return paragraph;
            }
        }

        public WordWatermark AddWatermark(WordWatermarkStyle watermarkStyle, string text) {
            // return new WordWatermark(this._document, this, this.Header.Default, watermarkStyle, text);
            return new WordWatermark(this._document, this, watermarkStyle, text);
        }

        public WordSection SetBorders(WordBorder wordBorder) {
            this.Borders.SetBorder(wordBorder);

            return this;
        }

        public WordParagraph AddHorizontalLine(BorderValues lineType = BorderValues.Single, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            return this.AddParagraph("").AddHorizontalLine(lineType, color, size, space);
        }

        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph("").AddHyperLink(text, uri, addStyle, tooltip, history);
        }

        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph("").AddHyperLink(text, anchor, addStyle, tooltip, history);
        }

        public void AddHeadersAndFooters() {
            WordHeadersAndFooters.AddHeadersAndFooters(this);
        }

        public WordTextBox AddTextBox(string text, WrapTextImage wrapTextImage = WrapTextImage.Square) {
            return AddParagraph(newRun: true).AddTextBox(text, wrapTextImage);
        }
    }
}
