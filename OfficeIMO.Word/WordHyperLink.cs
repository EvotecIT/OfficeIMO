using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

namespace OfficeIMO.Word {

    /// <summary>
    /// Defines hyperlink target frames.
    /// </summary>
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

    /// <summary>
    /// Represents a hyperlink element within a Word document.
    /// Provides helper methods for modifying hyperlink text,
    /// formatting, and target information.
    /// </summary>
    public class WordHyperLink : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly Hyperlink _hyperlink;

        /// <summary>
        /// Gets or sets the URI of the hyperlink.
        /// </summary>
        public System.Uri? Uri {
            get {
                var list = _document._wordprocessingDocument!.MainDocumentPart!.HyperlinkRelationships;
                foreach (var l in list) {
                    if (l.Id == _hyperlink.Id) {
                        return l.Uri;
                    }
                }

                return null;
            }
            set {
                if (value != null) {
                    var rel = _document._wordprocessingDocument!.MainDocumentPart!.AddHyperlinkRelationship(value, true);
                    _hyperlink.Id = rel.Id;
                }
            }
        }

        /// <summary>
        /// Specifies a location in the target of the hyperlink, in the case in which the link is an external link.
        /// </summary>
        public string? Id {
            get {
                return _hyperlink.Id;
            }
            set {
                _hyperlink.Id = value;
            }
        }

        internal Run? _run {
            get {
                return _hyperlink.Descendants<Run>().FirstOrDefault();
            }
        }

        internal RunProperties? _runProperties {
            get {
                return _hyperlink.Descendants<RunProperties>().FirstOrDefault();
            }
        }

        /// <summary>
        /// Indicates whether the hyperlink uses the mailto scheme.
        /// </summary>
        public bool IsEmail => Uri?.Scheme == Uri.UriSchemeMailto;

        /// <summary>
        /// Gets the email address if the hyperlink is a mailto link.
        /// </summary>
        public string EmailAddress {
            get {
                var uri = Uri;
                if (uri != null && IsEmail) {
                    return uri.AbsoluteUri.Replace(uri.PathAndQuery, "").Replace("mailto:", "");
                }

                return string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets whether the hyperlink is marked as visited.
        /// </summary>
        public bool History {
            get {
                return _hyperlink.History?.Value ?? false;
            }
            set {
                _hyperlink.History = value;
            }
        }

        /// <summary>
        /// Specifies a location in the target of the hyperlink, in the case in which the link is an external link.
        /// </summary>
        public string? DocLocation {
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
        public string? Anchor {
            get {
                return _hyperlink.Anchor;
            }
            set {
                _hyperlink.Anchor = value;
            }
        }

        /// <summary>
        /// Gets or sets the tooltip displayed when hovering over the hyperlink.
        /// </summary>
        public string? Tooltip {
            get {
                return _hyperlink.Tooltip;
            }
            set {
                _hyperlink.Tooltip = value;
            }
        }

        /// <summary>
        /// Gets or sets the target frame for the hyperlink.
        /// </summary>
        public TargetFrame? TargetFrame {
            get {
                string? target = _hyperlink.TargetFrame;
                if (!string.IsNullOrEmpty(target)) {
                    var targetFrame = (TargetFrame)Enum.Parse(typeof(TargetFrame), target, true);
                    return targetFrame;
                }

                return null;
            }
            set {
                _hyperlink.TargetFrame = value?.ToString();
            }
        }

        /// <summary>
        /// Gets a value indicating whether the hyperlink uses the HTTP or HTTPS scheme.
        /// </summary>
        public bool IsHttp => Uri != null && (Uri.Scheme == Uri.UriSchemeHttps || Uri.Scheme == Uri.UriSchemeHttp);

        /// <summary>
        /// Gets the scheme component of the hyperlink URI.
        /// </summary>
        public string? Scheme => Uri?.Scheme;

        /// <summary>
        /// Gets or sets the display text for the hyperlink.
        /// </summary>
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

        /// <summary>
        /// Initializes a new instance of the <see cref="WordHyperLink"/> class.
        /// </summary>
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
                OpenXmlElement? parent = _paragraph.Parent;
                while (parent != null && parent is not Body && parent is not Header && parent is not Footer) {
                    parent = parent.Parent;
                }

                OpenXmlPart? part = _document._wordprocessingDocument?.MainDocumentPart;
                var headerPart = (parent as Header)?.HeaderPart;
                var footerPart = (parent as Footer)?.FooterPart;

                if (headerPart != null) {
                    part = headerPart;
                } else if (footerPart != null) {
                    part = footerPart;
                }

                var rel = part?.HyperlinkRelationships?.FirstOrDefault(r => r.Id == _hyperlink.Id);
                if (rel != null) {
                    part!.DeleteReferenceRelationship(rel);
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

        /// <summary>
        /// Adds a hyperlink pointing to the specified anchor within the document.
        /// </summary>
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

        /// <summary>
        /// Adds a hyperlink pointing to an external URI.
        /// </summary>
        public static WordParagraph AddHyperLink(WordParagraph paragraph, string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            // Create a hyperlink relationship. Pass the relationship id to the hyperlink below.

            HyperlinkRelationship rel;

            // Determine if the paragraph belongs to a header or footer by checking the ancestors.
            var headerPart = paragraph._paragraph.Ancestors<Header>().FirstOrDefault()?.HeaderPart;
            var footerPart = paragraph._paragraph.Ancestors<Footer>().FirstOrDefault()?.FooterPart;

            if (headerPart != null) {
                rel = headerPart.AddHyperlinkRelationship(uri, true);
            } else if (footerPart != null) {
                rel = footerPart.AddHyperlinkRelationship(uri, true);
            } else {
                // Default to the main document part for paragraphs that are
                // located in the body or in elements such as text boxes or tables.
                rel = paragraph._document._wordprocessingDocument.MainDocumentPart!.AddHyperlinkRelationship(uri, true);
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

        /// <summary>
        /// Creates a hyperlink inserted after the reference link while copying
        /// the reference formatting.
        /// </summary>
        /// <param name="reference">Existing hyperlink used for formatting.</param>
        /// <param name="newText">Text for the new hyperlink.</param>
        /// <param name="newUri">Destination of the new hyperlink.</param>
        /// <returns>Newly created hyperlink.</returns>
        public static WordHyperLink CreateFormattedHyperlink(WordHyperLink reference, string newText, Uri newUri) {
            if (reference == null) throw new ArgumentNullException(nameof(reference));

            return reference.InsertFormattedHyperlinkAfter(newText, newUri);
        }

        /// <summary>
        /// Inserts a hyperlink after this hyperlink and copies this link's formatting.
        /// </summary>
        /// <param name="newText">Text for the new hyperlink.</param>
        /// <param name="newUri">Destination of the new hyperlink.</param>
        /// <returns>The inserted hyperlink.</returns>
        public WordHyperLink InsertFormattedHyperlinkAfter(string newText, Uri newUri) {
            if (newText == null) throw new ArgumentNullException(nameof(newText));
            if (newUri == null) throw new ArgumentNullException(nameof(newUri));

            HyperlinkRelationship rel;
            var headerPart = _paragraph.Ancestors<Header>().FirstOrDefault()?.HeaderPart;
            var footerPart = _paragraph.Ancestors<Footer>().FirstOrDefault()?.FooterPart;

            if (headerPart != null) {
                rel = headerPart.AddHyperlinkRelationship(newUri, true);
            } else if (footerPart != null) {
                rel = footerPart.AddHyperlinkRelationship(newUri, true);
            } else {
                rel = _document._wordprocessingDocument!.MainDocumentPart!.AddHyperlinkRelationship(newUri, true);
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

        /// <summary>
        /// Inserts a hyperlink before this hyperlink and copies this link's formatting.
        /// </summary>
        /// <param name="newText">Text for the new hyperlink.</param>
        /// <param name="newUri">Destination of the new hyperlink.</param>
        /// <returns>The inserted hyperlink.</returns>
        public WordHyperLink InsertFormattedHyperlinkBefore(string newText, Uri newUri) {
            if (newText == null) throw new ArgumentNullException(nameof(newText));
            if (newUri == null) throw new ArgumentNullException(nameof(newUri));

            HyperlinkRelationship rel;
            var headerPart = _paragraph.Ancestors<Header>().FirstOrDefault()?.HeaderPart;
            var footerPart = _paragraph.Ancestors<Footer>().FirstOrDefault()?.FooterPart;

            if (headerPart != null) {
                rel = headerPart.AddHyperlinkRelationship(newUri, true);
            } else if (footerPart != null) {
                rel = footerPart.AddHyperlinkRelationship(newUri, true);
            } else {
                rel = _document._wordprocessingDocument!.MainDocumentPart!.AddHyperlinkRelationship(newUri, true);
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

        /// <summary>
        /// Creates a copy of the given hyperlink and inserts it after the source link.
        /// </summary>
        /// <param name="reference">Hyperlink to duplicate.</param>
        /// <returns>The duplicated hyperlink.</returns>
        public static WordHyperLink DuplicateHyperlink(WordHyperLink reference) {
            if (reference == null) throw new ArgumentNullException(nameof(reference));

            Hyperlink duplicate = (Hyperlink)reference._hyperlink.CloneNode(true);
            reference._hyperlink.InsertAfterSelf(duplicate);

            return new WordHyperLink(reference._document, reference._paragraph, duplicate);
        }

        /// <summary>
        /// Copies run formatting from another hyperlink to this hyperlink.
        /// </summary>
        /// <param name="reference">Hyperlink to copy formatting from.</param>
        public void CopyFormattingFrom(WordHyperLink reference) {
            if (reference == null) throw new ArgumentNullException(nameof(reference));

            Run run = _run ?? new Run();
            if (_run == null) {
                _hyperlink.Append(run);
            }

            _runProperties?.Remove();

            if (reference._runProperties != null) {
                RunProperties clone = (RunProperties)reference._runProperties.CloneNode(true);
                run.PrependChild(clone);
            }
        }
    }
}