using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordFooter {
        public readonly List<WordParagraph> Paragraphs = new List<WordParagraph>();
        private readonly string _type;
        private readonly Footer _footerFirst;
        private readonly Footer _footerOdd;
        private readonly Footer _footerEven;
        
        internal WordFooter(WordDocument document, string type) {
            if (type == "first") {
                _footerFirst = document._footerFirst;
            } else if (type == "odd") {
                _footerOdd = document._footerOdd;
            } else if (type == "even") {
                _footerEven = document._footerEven;
            }
            _type = type;
        }
        public WordParagraph InsertParagraph() {
            var wordParagraph = new WordParagraph();
            if (_type == "first") {
                _footerFirst.Append(wordParagraph._paragraph);
            } else if (_type == "odd") {
                _footerOdd.Append(wordParagraph._paragraph);
            } else if (_type == "even") {
                _footerEven.Append(wordParagraph._paragraph);
            }
            this.Paragraphs.Add(wordParagraph);
            return wordParagraph;
        }
    }
}
