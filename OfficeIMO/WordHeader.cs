using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordHeader {
        public readonly List<WordParagraph> Paragraphs = new List<WordParagraph>();
        private readonly HeaderFooterValues _type;
        private readonly Header _headerFirst;
        private readonly Header _headerDefault;
        private readonly Header _headerEven;

        internal WordHeader(WordDocument document, HeaderFooterValues type, Header headerPartHeader) {
            if (type == HeaderFooterValues.First) {
                _headerFirst = headerPartHeader;
            } else if (type == HeaderFooterValues.Default) {
                _headerDefault = headerPartHeader;
            } else if (type == HeaderFooterValues.Even) {
                _headerEven = headerPartHeader;
            }
            _type = type;
        }
        public WordParagraph InsertParagraph() {
            var wordParagraph = new WordParagraph();
            if (_type == HeaderFooterValues.First) {
                _headerFirst.Append(wordParagraph._paragraph);
            } else if (_type == HeaderFooterValues.Default) {
                _headerDefault.Append(wordParagraph._paragraph);
            } else if (_type == HeaderFooterValues.Even) {
                _headerEven.Append(wordParagraph._paragraph);
            }
            this.Paragraphs.Add(wordParagraph);
            return wordParagraph;
        }
    }
}
