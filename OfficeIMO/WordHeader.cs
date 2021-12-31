using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordHeader {
        public readonly List<WordParagraph> Paragraphs = new List<WordParagraph>();
        private readonly string _type;
        private readonly Header _headerFirst;
        private readonly Header _headerOdd;
        private readonly Header _headerEven;

        internal WordHeader(WordDocument document, string type) {
            if (type == "first") {
                _headerFirst = document._headerFirst;
            } else if (type == "odd") {
                _headerOdd = document._headerOdd;
            } else if (type == "even") {
                _headerEven = document._headerEven;
            }
            _type = type;
        }
        public WordParagraph InsertParagraph() {
            var wordParagraph = new WordParagraph();
            if (_type == "first") {
                _headerFirst.Append(wordParagraph._paragraph);
            } else if (_type == "odd") {
                _headerOdd.Append(wordParagraph._paragraph);
            } else if (_type == "even") {
                _headerEven.Append(wordParagraph._paragraph);
            }
            this.Paragraphs.Add(wordParagraph);
            return wordParagraph;
        }
    }
}
