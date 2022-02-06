using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordHyperLink {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly Hyperlink _hyperlink;

        public readonly System.Uri Url;

        public List<WordParagraph> Paragraphs {
            get {
                var list = new List<WordParagraph>();
                return list;
                // return WordSection.ConvertParagraphsToWordParagraphs(_document, _hyperlink.ChildElements.Of);
            }
        }

        public bool IsEmail => Url.Scheme == Uri.UriSchemeMailto;

        public string EmailAddress {
            get {
                if (IsEmail) {
                    return Url.AbsoluteUri.Replace(Url.PathAndQuery, "").Replace("mailto:", "");
                }

                return "";
            }
        }

        public bool IsHttp => Url.Scheme == Uri.UriSchemeHttps || Url.Scheme == Uri.UriSchemeHttp;

        public string Scheme => Url.Scheme;

        public WordHyperLink(WordDocument document, Paragraph paragraph, Hyperlink hyperlink) {
            _document = document;
            _paragraph = paragraph;
            _hyperlink = hyperlink;

            var list = document._wordprocessingDocument.MainDocumentPart.HyperlinkRelationships;
            foreach (var l in list) {
                if (l.Id == _hyperlink.Id) {
                    Url = l.Uri;
                }
            }
        }
    }
}
