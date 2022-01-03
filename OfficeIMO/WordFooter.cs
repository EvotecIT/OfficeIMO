using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordFooter {
        public readonly List<WordParagraph> Paragraphs = new List<WordParagraph>();
        private readonly HeaderFooterValues _type;
        private readonly Footer _footerFirst;
        private readonly Footer _footerDefault;
        private readonly Footer _footerEven;
        
        internal WordFooter(WordDocument document, HeaderFooterValues type, Footer footerPartFooter) {
            if (type == HeaderFooterValues.First) {
                _footerFirst = footerPartFooter;
            } else if (type == HeaderFooterValues.Default) {
                _footerDefault = footerPartFooter;
            } else if (type == HeaderFooterValues.Even) {
                _footerEven = footerPartFooter;
            }
            _type = type;
        }
        public WordParagraph InsertParagraph() {
            var wordParagraph = new WordParagraph();
            if (_type == HeaderFooterValues.First) {
                _footerFirst.Append(wordParagraph._paragraph);
            } else if (_type == HeaderFooterValues.Default) {
                _footerDefault.Append(wordParagraph._paragraph);
            } else if (_type == HeaderFooterValues.Even) {
                _footerEven.Append(wordParagraph._paragraph);
            }
            this.Paragraphs.Add(wordParagraph);
            return wordParagraph;
        }
    }
}
