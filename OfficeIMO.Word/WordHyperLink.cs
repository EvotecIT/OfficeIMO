using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    public enum TargetFrame {
        /// <summary>
        /// opens in the current window
        /// </summary>
        _top,
        /// <summary>
        /// Opens in the current window
        /// </summary>
        _self,
        /// <summary>
        /// opens in the parent of the current frame
        /// </summary>
        _parent,
        /// <summary>
        /// opens in a new web browser window
        /// </summary>
        _blank
    }

    public class WordHyperLink : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly Hyperlink _hyperlink;

        public System.Uri Uri {
            get {
                var list = _document._wordprocessingDocument.MainDocumentPart.HyperlinkRelationships;
                foreach (var l in list) {
                    if (l.Id == _hyperlink.Id) {
                        return l.Uri;
                    }
                }

                return null;
            }
            set {
                var rel = _document._wordprocessingDocument.MainDocumentPart.AddHyperlinkRelationship(value, true);
                if (rel != null) {
                    _hyperlink.Id = rel.Id;
                }
            }
        }

        /// <summary>
        /// Specifies a location in the target of the hyperlink, in the case in which the link is an external link.
        /// </summary>
        public string Id {
            get {
                return _hyperlink.Id;
            }
            set {
                _hyperlink.Id = value;
            }
        }

        internal Run _run {
            get {
                return _hyperlink.Descendants<Run>().FirstOrDefault();
            }
        }

        internal RunProperties _runProperties {
            get {
                return _hyperlink.Descendants<RunProperties>().FirstOrDefault();
            }
        }

        public bool IsEmail => Uri.Scheme == Uri.UriSchemeMailto;

        public string EmailAddress {
            get {
                if (IsEmail) {
                    return Uri.AbsoluteUri.Replace(Uri.PathAndQuery, "").Replace("mailto:", "");
                }

                return "";
            }
        }

        public bool History {
            get {
                return _hyperlink.History;
            }
            set {
                _hyperlink.History = value;
            }
        }

        /// <summary>
        /// Specifies a location in the target of the hyperlink, in the case in which the link is an external link.
        /// </summary>
        public string DocLocation {
            get {
                return _hyperlink.DocLocation;
            }
            set {
                _hyperlink.DocLocation = value;
            }
        }

        /// <summary>
        /// Specifies the name of a bookmark within the document.
        /// See Bookmark. If the attribute is omitted, then the default behavior is to navigate to the start of the document.
        /// If the r:id attribute is specified, then the anchor attribute is ignored.
        /// </summary>
        public string Anchor {
            get {
                return _hyperlink.Anchor;
            }
            set {
                _hyperlink.Anchor = value;
            }
        }

        public string Tooltip {
            get {
                return _hyperlink.Tooltip;
            }
            set {
                _hyperlink.Tooltip = value;
            }
        }

        public TargetFrame? TargetFrame {
            get {
                if (_hyperlink != null) {
                    string target = _hyperlink.TargetFrame;
                    if (target != null) {
                        var targetFrame = (TargetFrame)Enum.Parse(typeof(TargetFrame), target, true);
                        return targetFrame;
                    }
                }

                return null;
            }
            set {
                _hyperlink.TargetFrame = value.ToString();
            }
        }

        public bool IsHttp => Uri.Scheme == Uri.UriSchemeHttps || Uri.Scheme == Uri.UriSchemeHttp;

        public string Scheme => Uri.Scheme;

        public string Text {
            get {
                var run = _hyperlink.ChildElements.OfType<Run>().FirstOrDefault();
                if (run != null) {
                    var text = run.ChildElements.OfType<Text>().FirstOrDefault();
                    if (text != null) {
                        return text.Text;
                    }
                }
                return "";
            }
            set {
                var run = _hyperlink.ChildElements.OfType<Run>().FirstOrDefault();
                if (run != null) {
                    var text = run.ChildElements.OfType<Text>().FirstOrDefault();
                    if (text != null) {
                        if (value != null) {
                            text.Text = value;
                        } else {
                            text.Remove();
                            run.Remove();
                        }
                    }
                }
            }
        }

        public WordHyperLink(WordDocument document, Paragraph paragraph, Hyperlink hyperlink) {
            _document = document;
            _paragraph = paragraph;
            _hyperlink = hyperlink;
        }

        /// <summary>
        /// Removes hyperlink. When specified to remove paragraph it will only do so,
        /// if paragraph is empty or contains only paragraph properties.
        /// </summary>
        /// <param name="includingParagraph"></param>
        public void Remove(bool includingParagraph = true) {
            RemoveHyperLink(includingParagraph);
        }

        /// <summary>
        /// Removes hyperlink and detaches related relationship. When specified
        /// to remove paragraph it will only do so if paragraph is empty or
        /// contains only paragraph properties.
        /// </summary>
        /// <param name="includingParagraph"></param>
        public void RemoveHyperLink(bool includingParagraph = true) {
            if (!string.IsNullOrEmpty(_hyperlink.Id)) {
                OpenXmlElement parent = _paragraph.Parent;
                while (parent != null && !(parent is Body) && !(parent is Header) && !(parent is Footer)) {
                    parent = parent.Parent;
                }

                OpenXmlPart part = _document._wordprocessingDocument.MainDocumentPart;
                if (parent is Header header) {
                    part = header.HeaderPart;
                } else if (parent is Footer footer) {
                    part = footer.FooterPart;
                }

                var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == _hyperlink.Id);
                if (rel != null) {
                    part.DeleteReferenceRelationship(rel);
                }
            }

            _hyperlink.Remove();
            if (includingParagraph) {
                if (this._paragraph.ChildElements.Count == 0) {
                    this._paragraph.Remove();
                } else if (this._paragraph.ChildElements.Count == 1 && this._paragraph.ChildElements.OfType<ParagraphProperties>().Any()) {
                    this._paragraph.Remove();
                }
            }
        }

        public static WordParagraph AddHyperLink(WordParagraph paragraph, string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            Hyperlink hyperlink = new Hyperlink() {
                Anchor = anchor,
                //DocLocation = "",
                History = history,
            };

            Run run = new Run(new Text(text) {
                Space = SpaceProcessingModeValues.Preserve
            });

            // Styling for the hyperlink
            if (addStyle) {
                RunProperties runPropertiesHyperLink = new RunProperties(
                    new RunStyle { Val = "Hyperlink", },
                    new Color { ThemeColor = ThemeColorValues.Hyperlink, Val = "0000FF" },
                    new Underline { Val = UnderlineValues.Single }
                );
                run.RunProperties = runPropertiesHyperLink;
            }

            if (tooltip != "") {
                hyperlink.Tooltip = tooltip;
            }

            hyperlink.Append(run);
            paragraph._paragraph.Append(hyperlink);
            paragraph._hyperlink = hyperlink;
            return paragraph;
        }

        public static WordParagraph AddHyperLink(WordParagraph paragraph, string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            // Create a hyperlink relationship. Pass the relationship id to the hyperlink below.

            HyperlinkRelationship rel;

            // Determine if the paragraph belongs to a header or footer by checking the ancestors.
            var header = paragraph._paragraph.Ancestors<Header>().FirstOrDefault();
            var footer = paragraph._paragraph.Ancestors<Footer>().FirstOrDefault();

            if (header != null) {
                rel = header.HeaderPart.AddHyperlinkRelationship(uri, true);
            } else if (footer != null) {
                rel = footer.FooterPart.AddHyperlinkRelationship(uri, true);
            } else {
                // Default to the main document part for paragraphs that are
                // located in the body or in elements such as text boxes or tables.
                rel = paragraph._document._wordprocessingDocument.MainDocumentPart.AddHyperlinkRelationship(uri, true);
            }

            Hyperlink hyperlink = new Hyperlink() {
                Id = rel.Id,
                //DocLocation = "",
                History = history,
            };

            Run run = new Run(new Text(text) {
                Space = SpaceProcessingModeValues.Preserve
            });

            // Styling for the hyperlink
            if (addStyle) {
                RunProperties runPropertiesHyperLink = new RunProperties(
                    new RunStyle { Val = "Hyperlink", },
                    new Color { ThemeColor = ThemeColorValues.Hyperlink, Val = "0000FF" },
                    new Underline { Val = UnderlineValues.Single }
                );
                run.RunProperties = runPropertiesHyperLink;
            }

            if (tooltip != "") {
                hyperlink.Tooltip = tooltip;
            }

            hyperlink.Append(run);
            paragraph._paragraph.Append(hyperlink);
            paragraph._hyperlink = hyperlink;
            return paragraph;
        }

        public static WordHyperLink CreateFormattedHyperlink(WordHyperLink reference, string newText, Uri newUri) {
            if (reference == null) throw new ArgumentNullException(nameof(reference));

            return reference.InsertFormattedHyperlinkAfter(newText, newUri);
        }

        public WordHyperLink InsertFormattedHyperlinkAfter(string newText, Uri newUri) {
            if (newText == null) throw new ArgumentNullException(nameof(newText));
            if (newUri == null) throw new ArgumentNullException(nameof(newUri));

            HyperlinkRelationship rel;
            var header = _paragraph.Ancestors<Header>().FirstOrDefault();
            var footer = _paragraph.Ancestors<Footer>().FirstOrDefault();

            if (header != null) {
                rel = header.HeaderPart.AddHyperlinkRelationship(newUri, true);
            } else if (footer != null) {
                rel = footer.FooterPart.AddHyperlinkRelationship(newUri, true);
            } else {
                rel = _document._wordprocessingDocument.MainDocumentPart.AddHyperlinkRelationship(newUri, true);
            }

            Hyperlink hyperlink = new Hyperlink() {
                Id = rel.Id,
                History = _hyperlink.History
            };

            Run run = new Run(new Text(newText) {
                Space = SpaceProcessingModeValues.Preserve
            });

            if (_runProperties != null) {
                run.RunProperties = (RunProperties)_runProperties.CloneNode(true);
            }

            hyperlink.Append(run);

            _hyperlink.InsertAfterSelf(hyperlink);

            return new WordHyperLink(_document, _paragraph, hyperlink);
        }

        public WordHyperLink InsertFormattedHyperlinkBefore(string newText, Uri newUri) {
            if (newText == null) throw new ArgumentNullException(nameof(newText));
            if (newUri == null) throw new ArgumentNullException(nameof(newUri));

            HyperlinkRelationship rel;
            var header = _paragraph.Ancestors<Header>().FirstOrDefault();
            var footer = _paragraph.Ancestors<Footer>().FirstOrDefault();

            if (header != null) {
                rel = header.HeaderPart.AddHyperlinkRelationship(newUri, true);
            } else if (footer != null) {
                rel = footer.FooterPart.AddHyperlinkRelationship(newUri, true);
            } else {
                rel = _document._wordprocessingDocument.MainDocumentPart.AddHyperlinkRelationship(newUri, true);
            }

            Hyperlink hyperlink = new Hyperlink() {
                Id = rel.Id,
                History = _hyperlink.History
            };

            Run run = new Run(new Text(newText) {
                Space = SpaceProcessingModeValues.Preserve
            });

            if (_runProperties != null) {
                run.RunProperties = (RunProperties)_runProperties.CloneNode(true);
            }

            hyperlink.Append(run);

            _hyperlink.InsertBeforeSelf(hyperlink);

            return new WordHyperLink(_document, _paragraph, hyperlink);
        }

        public static WordHyperLink DuplicateHyperlink(WordHyperLink reference) {
            if (reference == null) throw new ArgumentNullException(nameof(reference));

            Hyperlink duplicate = (Hyperlink)reference._hyperlink.CloneNode(true);
            reference._hyperlink.InsertAfterSelf(duplicate);

            return new WordHyperLink(reference._document, reference._paragraph, duplicate);
        }
    }
}
