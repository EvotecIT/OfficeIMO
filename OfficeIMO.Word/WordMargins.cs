using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordMargins {
        private readonly WordDocument _document;
        private readonly WordSection _section;

        public WordMargins(WordDocument wordDocument, WordSection wordSection) {
            _document = wordDocument;
            _section = wordSection;
        }
        public UInt32Value MarginLeft {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin != null) {
                    return pageMargin.Left;
                }

                return null;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(PageMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin.Left = value;
            }
        }
        public UInt32Value MarginRight {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin != null) {
                    return pageMargin.Right;
                }

                return null;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(PageMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin.Right = value;
            }
        }
        public int? MarginTop {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin != null) {
                    return pageMargin.Top;
                }

                return null;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(PageMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin.Top = value;
            }
        }
        public int? MarginBottom {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin != null) {
                    return pageMargin.Bottom;
                }

                return null;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(PageMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin.Bottom = value;
            }
        }

        public UInt32Value HeaderDistance {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin != null) {
                    return pageMargin.Header;
                }

                return null;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(PageMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin.Header = value;
            }
        }

        public UInt32Value FooterDistance {
            get {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin != null) {
                    return pageMargin.Footer;
                }

                return null;
            }
            set {
                var pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                if (pageMargin == null) {
                    _section._sectionProperties.Append(PageMargins.Normal);
                    pageMargin = _section._sectionProperties.GetFirstChild<PageMargin>();
                }

                pageMargin.Footer = value;
            }
        }
    }
}
