using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Reflection;
using System.Linq;
using System.Xml.Linq;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word {
    /// <summary>
    /// Contains public methods for editing paragraphs.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Add Break to the paragraph. By default it adds soft break (SHIFT+ENTER)
        /// </summary>
        /// <param name="breakType">Optional argument to add a specific type of break.</param>
        /// <returns>The new WordParagraph that this method creates.</returns>
        public WordParagraph AddBreak(BreakValues? breakType = null) {
            WordParagraph wordParagraph = new WordParagraph(this._document, this._paragraph, new Run());
            if (breakType != null) {
                this._paragraph.Append(new Run(new Break() { Type = breakType }));
            } else {
                this._paragraph.Append(new Run(new Break()));
            }
            return wordParagraph;
        }

        /// <summary>
        /// Remove the paragraph from WordDocument
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        public void Remove() {
            _list?.RemoveItem(this);
            if (_paragraph != null) {
                if (this._paragraph.Parent != null) {
                    if (this.IsBookmark) {
                        this.Bookmark!.Remove();
                    }

                    if (this.IsBreak) {
                        this.Break!.Remove();
                        // Removing a break can also remove the entire paragraph.
                        // When that happens there's nothing else to clean up.
                        if (this._paragraph.Parent == null) {
                            return;
                        }
                    }

                    // break should cover this
                    //if (this.IsPageBreak) {
                    //    this.PageBreak.Remove();
                    //}

                    if (this.IsEquation) {
                        this.Equation!.Remove();
                    }

                    if (this.IsHyperLink) {
                        this.RemoveHyperLink();
                    }

                    if (this.IsListItem) {

                    }

                    if (this.IsImage) {
                        this.Image!.Remove();
                    }

                    if (this.IsStructuredDocumentTag) {
                        this.StructuredDocumentTag!.Remove();
                    }

                    if (this.IsField) {
                        this.Field!.Remove();
                    }

                    var runs = this._paragraph.ChildElements.OfType<Run>().ToList();
                    if (runs.Count == 0) {
                        this._paragraph.Remove();
                    } else if (runs.Count == 1) {
                        this._paragraph.Remove();
                    } else {
                        foreach (var run in runs) {
                            if (run == _run) {
                                run.Remove();
                            }
                        }
                    }
                } else {
                    throw new InvalidOperationException($"Cannot remove paragraph because it no longer has a parent. Paragraph text: '{Text}'.");
                }
            } else {
                // this shouldn't happen
                throw new InvalidOperationException($"Cannot remove paragraph because it is not initialized in the document. Paragraph text: '{Text}'.");
            }
        }

        /// <summary>
        /// Add paragraph right after existing paragraph.
        /// This can be useful to add empty lines, or moving cursor to next line
        /// </summary>
        /// <param name="wordParagraph">Optional WordParagraph to insert after the
        /// given WordParagraph instead of at the end of the document.</param>
        /// <returns>The inserted WordParagraph.</returns>
        public WordParagraph AddParagraph(WordParagraph? wordParagraph = null) {
            if (wordParagraph is null) {
                // we create paragraph (and within that add it to document)
                wordParagraph = new WordParagraph(this._document, newParagraph: true, newRun: false);
            } else {
                EnsureParagraphCanBeInserted(this._document, this._paragraph, wordParagraph,
                    "insert a paragraph after the current paragraph", this);
            }

            this._paragraph.InsertAfterSelf(wordParagraph._paragraph);
            wordParagraph.RefreshParent();
            return wordParagraph;
        }

        /// <summary>
        /// Add a paragraph with the given text to the end of the document.
        /// </summary>
        /// <param name="text">The text to fill the paragraph with.</param>
        /// <returns> The appended WordParagraph.</returns>
        public WordParagraph AddParagraph(string text) {
            // we create paragraph (and within that add it to document)
            var wordParagraph = new WordParagraph(this._document, newParagraph: true, newRun: false) {
                Text = text
            };
            this._paragraph.InsertAfterSelf(wordParagraph._paragraph);
            wordParagraph.RefreshParent();
            return wordParagraph;
        }

        /// <summary>
        /// Insert a new paragraph after the WordParagraph AddParagraphAfterSelf is called on in the document.
        /// </summary>
        /// <returns>The inserted WordParagraph</returns>
        public WordParagraph AddParagraphAfterSelf() {
            WordParagraph paragraph = new WordParagraph(this._document, true, false);
            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            paragraph.RefreshParent();
            return paragraph;
        }

        /// <summary>
        /// Add paragraph after self but by allowing to specify section
        /// </summary>
        /// <param name="section">The section to add the paragraph to. When paragraph is given this has no effect.</param>
        /// <param name="paragraph">The optional paragraph to add the paragraph to.</param>
        /// <returns>The new WordParagraph</returns>
        public WordParagraph AddParagraphAfterSelf(WordSection section, WordParagraph? paragraph = null) {
            if (!ReferenceEquals(section._document, this._document)) {
                throw new InvalidOperationException("Cannot add a paragraph using a section from a different document.");
            }

            var owningSection = GetSectionForParagraph(this._document, this);
            if (owningSection != null) {
                if (!ReferenceEquals(owningSection, section)) {
                    throw new InvalidOperationException("The provided section does not match the section of the current paragraph.");
                }
            } else {
                var currentSection = GetSectionPropertiesForElement(this._paragraph);
                if (currentSection != null && !AreSectionsEquivalent(currentSection, section._sectionProperties)) {
                    throw new InvalidOperationException("The provided section does not match the section of the current paragraph.");
                }
            }

            if (paragraph is null) {
                paragraph = new WordParagraph(section._document, true, false);
            } else {
                EnsureParagraphCanBeInserted(this._document, this._paragraph, paragraph,
                    "insert a paragraph after the current paragraph", this);
            }

            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            paragraph.RefreshParent();
            return paragraph;
        }

        internal static void EnsureParagraphCanBeInserted(WordDocument document, OpenXmlElement anchorElement, WordParagraph candidate, string operationDescription, WordParagraph? anchorParagraph = null) {
            if (candidate is null) {
                throw new ArgumentNullException(nameof(candidate));
            }

            if (candidate._document != null && !ReferenceEquals(candidate._document, document)) {
                throw new InvalidOperationException($"Cannot {operationDescription} because the supplied paragraph belongs to a different document. Clone the paragraph or create it within the target document instead.");
            }

            if (candidate._document == null) {
                candidate._document = document;
                candidate.InvalidateParent();
            }

            if (candidate._paragraph.Parent == null) {
                return;
            }

            var anchorContainer = GetTopLevelContainer(anchorElement);
            var candidateContainer = GetTopLevelContainer(candidate._paragraph);

            if (!ReferenceEquals(anchorContainer, candidateContainer)) {
                throw new InvalidOperationException($"Cannot {operationDescription} because the supplied paragraph resides in '{DescribeContainer(candidateContainer)}' while the target is in '{DescribeContainer(anchorContainer)}'. Provide a paragraph from the same section or clone it first.");
            }

            if (anchorContainer is Body) {
                WordSection? anchorSection = anchorParagraph != null
                    ? GetSectionForParagraph(document, anchorParagraph)
                    : document.Sections.LastOrDefault();
                WordSection? candidateSection = GetSectionForParagraph(document, candidate);

                if (anchorSection != null && candidateSection != null) {
                    if (!ReferenceEquals(anchorSection, candidateSection)) {
                        throw new InvalidOperationException($"Cannot {operationDescription} because the supplied paragraph originates from a different section. Provide a paragraph from the same section or clone it first.");
                    }
                    return;
                }

                var anchorSectionProps = GetSectionPropertiesForElement(anchorElement);
                var candidateSectionProps = GetSectionPropertiesForElement(candidate._paragraph);
                if (!AreSectionsEquivalent(anchorSectionProps, candidateSectionProps)) {
                    throw new InvalidOperationException($"Cannot {operationDescription} because the supplied paragraph originates from a different section. Provide a paragraph from the same section or clone it first.");
                }
            }

            candidate.InvalidateParent();
        }

        private static SectionProperties? GetSectionPropertiesForElement(OpenXmlElement element) {
            var topLevel = GetTopLevelContainer(element);
            if (topLevel is null) {
                return null;
            }

            if (topLevel is Body bodyElement) {
                OpenXmlElement? bodyChild = GetBodyChildContainer(element);
                if (bodyChild?.Parent is Body owningBody) {
                    return GetSectionPropertiesForBodyChild(owningBody, bodyChild);
                }

                return GetLastSectionProperties(bodyElement);
            }

            if (topLevel.Parent is Body body) {
                return GetSectionPropertiesForBodyChild(body, topLevel);
            }

            return null;
        }

        private static SectionProperties? GetSectionPropertiesForBodyChild(Body body, OpenXmlElement bodyChild) {
            SectionProperties? ownBoundary = GetSectionBoundaryProperties(bodyChild);
            if (ownBoundary != null) {
                return ownBoundary;
            }

            bool foundElement = false;
            foreach (var child in body.ChildElements) {
                if (ReferenceEquals(child, bodyChild)) {
                    foundElement = true;
                    continue;
                }

                if (foundElement) {
                    SectionProperties? nextBoundary = GetSectionBoundaryProperties(child);
                    if (nextBoundary != null) {
                        return nextBoundary;
                    }
                }
            }

            return body.Elements<SectionProperties>().LastOrDefault();
        }

        private static SectionProperties? GetSectionBoundaryProperties(OpenXmlElement element) {
            return element switch {
                Paragraph paragraph => paragraph.ParagraphProperties?.SectionProperties,
                SectionProperties sectionProperties => sectionProperties,
                _ => null
            };
        }

        private static SectionProperties? GetLastSectionProperties(Body body) {
            SectionProperties? last = null;
            foreach (var child in body.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        var sp = paragraph.ParagraphProperties?.SectionProperties;
                        if (sp != null) {
                            last = sp;
                        }
                        break;
                    case SectionProperties sectionProperties:
                        last = sectionProperties;
                        break;
                }
            }

            return last ?? body.Elements<SectionProperties>().LastOrDefault();
        }

        private static OpenXmlElement? GetTopLevelContainer(OpenXmlElement element) {
            OpenXmlElement? current = element;
            while (current.Parent != null && current is not Body && current is not Header && current is not Footer) {
                current = current.Parent;
            }
            return current;
        }

        private static bool AreSectionsEquivalent(SectionProperties? left, SectionProperties? right) {
            if (left == null || right == null) {
                return left == right;
            }

            if (ReferenceEquals(left, right)) {
                return true;
            }

            return string.Equals(left.OuterXml, right.OuterXml, StringComparison.Ordinal);
        }

        private static string DescribeContainer(OpenXmlElement? container) {
            if (container is null) {
                return "an unknown container";
            }

            return container switch {
                Body => "the document body",
                Header => "a document header",
                Footer => "a document footer",
                _ => container.GetType().Name
            };
        }

        private static WordSection? GetSectionForParagraph(WordDocument document, WordParagraph paragraph) {
            foreach (var section in document.Sections) {
                if (section.Paragraphs.Any(p => ReferenceEquals(p._paragraph, paragraph._paragraph))) {
                    return section;
                }
            }

            return null;
        }

        /// <summary>
        /// Add a paragraph before the paragraph that AddParagraphBeforeSelf was called on.
        /// </summary>
        /// <returns>The inserted paragraph</returns>
        public WordParagraph AddParagraphBeforeSelf() {
            WordParagraph paragraph = new WordParagraph(this._document, true, false);
            this._paragraph.InsertBeforeSelf(paragraph._paragraph);
            //document.Paragraphs.Add(paragraph);
            paragraph.RefreshParent();
            return paragraph;
        }
    }
}
