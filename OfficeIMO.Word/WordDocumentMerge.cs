using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Imports the body of one Word package into another, reconciling styles, numbering, stories and identifiers.
    /// </summary>
    internal sealed partial class WordDocumentMerge {
        private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private readonly WordDocument _destination;
        private readonly MainDocumentPart _sourceMain;
        private readonly MainDocumentPart _destinationMain;
        private readonly WordDocumentMergeOptions _options;
        private readonly Dictionary<OpenXmlPart, OpenXmlPart> _partMap = new();
        private readonly Dictionary<(OpenXmlPartContainer Source, OpenXmlPartContainer Destination), Dictionary<string, string>> _relationshipMaps = new();

        internal WordDocumentMerge(WordDocument destination, MainDocumentPart sourceMain, MainDocumentPart destinationMain, WordDocumentMergeOptions options) {
            _destination = destination;
            _sourceMain = sourceMain;
            _destinationMain = destinationMain;
            _options = options;
        }

        internal void Append(Body sourceBody, Body destinationBody) {
            // Identifiers already in use are captured before anything is imported.
            var identifiers = new DestinationIdentifiers(_destinationMain);

            var staging = new Body();
            foreach (OpenXmlElement element in sourceBody.ChildElements) {
                if (element is not SectionProperties) {
                    staging.Append(element.CloneNode(true));
                }
            }

            var sourceSection = (SectionProperties?)sourceBody.Elements<SectionProperties>().LastOrDefault()?.CloneNode(true);
            if (_options.SectionMode == WordMergeSectionMode.UseDestinationSections) {
                RemoveSectionBreaks(staging);
                sourceSection = null;
            }

            PlanStyles();
            ImportNumbering();
            CommitStyles();

            List<OpenXmlElement> notes = ImportNotes(staging);
            List<OpenXmlElement> comments = ImportComments(new OpenXmlElement[] { staging }.Concat(notes));

            RemapRelationshipIds(staging, _sourceMain, _destinationMain);
            if (sourceSection != null) {
                RemapRelationshipIds(sourceSection, _sourceMain, _destinationMain);
            }

            List<OpenXmlElement> stories = GetImportedHeaderFooterRoots();
            var roots = new List<OpenXmlElement> { staging };
            roots.AddRange(notes);
            roots.AddRange(comments);
            roots.AddRange(stories);

            foreach (OpenXmlElement root in roots) {
                RemapStyleReferences(root);
                RemapNumberingReferences(root);
                _destination.AssignNewSdtIds(root);
                WordDrawingIdAllocator.Reassign(_destination, root);
            }

            ApplyImplicitDefaultStyles(new OpenXmlElement[] { staging }.Concat(notes).Concat(comments));
            RemapBookmarks(roots, identifiers);
            Dictionary<string, string> paragraphIds = RemapParagraphIds(roots, identifiers);
            ImportCommentExtensions(comments, paragraphIds);
            MergeImportedNamespaces();

            PlaceContent(destinationBody, staging, sourceSection);
        }

        /// <summary>
        /// Imported elements can carry attributes from newer Word namespaces (w14, w15, ...) that the source part
        /// declared as ignorable. The destination parts receive the same declarations so the markup stays valid.
        /// </summary>
        private void MergeImportedNamespaces() {
            MergeNamespaces(_sourceMain.Document, _destinationMain.Document);
            MergeNamespaces(_sourceMain.StyleDefinitionsPart?.Styles, _destinationMain.StyleDefinitionsPart?.Styles);
            MergeNamespaces(_sourceMain.NumberingDefinitionsPart?.Numbering, _destinationMain.NumberingDefinitionsPart?.Numbering);
            MergeNamespaces(_sourceMain.FootnotesPart?.Footnotes, _destinationMain.FootnotesPart?.Footnotes);
            MergeNamespaces(_sourceMain.EndnotesPart?.Endnotes, _destinationMain.EndnotesPart?.Endnotes);
            MergeNamespaces(_sourceMain.WordprocessingCommentsPart?.Comments, _destinationMain.WordprocessingCommentsPart?.Comments);
            MergeNamespaces(_sourceMain.WordprocessingCommentsExPart?.CommentsEx, _destinationMain.WordprocessingCommentsExPart?.CommentsEx);
        }

        private static void MergeNamespaces(OpenXmlPartRootElement? source, OpenXmlPartRootElement? destination) {
            if (source == null || destination == null) {
                return;
            }

            var declared = destination.NamespaceDeclarations.ToDictionary(_ => _.Key, _ => _.Value, StringComparer.Ordinal);
            foreach (KeyValuePair<string, string> declaration in source.NamespaceDeclarations) {
                if (!declared.ContainsKey(declaration.Key)) {
                    destination.AddNamespaceDeclaration(declaration.Key, declaration.Value);
                    declared.Add(declaration.Key, declaration.Value);
                }
            }

            string? sourceIgnorable = source.MCAttributes?.Ignorable?.Value;
            if (string.IsNullOrWhiteSpace(sourceIgnorable)) {
                return;
            }

            var ignorable = (destination.MCAttributes?.Ignorable?.Value ?? string.Empty)
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .ToList();
            bool changed = false;
            foreach (string prefix in sourceIgnorable!.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (!ignorable.Contains(prefix) && declared.ContainsKey(prefix)) {
                    ignorable.Add(prefix);
                    changed = true;
                }
            }

            if (changed) {
                destination.MCAttributes ??= new MarkupCompatibilityAttributes();
                destination.MCAttributes.Ignorable = string.Join(" ", ignorable);
            }
        }

        private static void RemoveSectionBreaks(OpenXmlElement root) {
            foreach (SectionProperties sectionProperties in root.Descendants<SectionProperties>().ToList()) {
                if (sectionProperties.Parent is ParagraphProperties) {
                    sectionProperties.Remove();
                }
            }
        }

        /// <summary>
        /// Appends the staged content so every section keeps its own page setup: the destination's closing
        /// section properties become a paragraph-level section break and the source's become the new closing ones.
        /// </summary>
        private void PlaceContent(Body destinationBody, Body staging, SectionProperties? sourceSection) {
            SectionProperties? destinationSection = destinationBody.Elements<SectionProperties>().LastOrDefault();
            destinationSection?.Remove();
            bool destinationHasContent = destinationBody.ChildElements.Any(_ => _ is not SectionProperties);

            SectionProperties? closingSection = destinationSection;
            if (sourceSection != null) {
                if (destinationSection != null && destinationHasContent) {
                    AttachSectionBreak(destinationBody, destinationSection);
                }

                closingSection = sourceSection;
            }

            foreach (OpenXmlElement element in staging.ChildElements.ToList()) {
                element.Remove();
                destinationBody.Append(element);
            }

            if (closingSection != null) {
                destinationBody.Append(closingSection);
            }
        }

        private static void AttachSectionBreak(Body body, SectionProperties sectionProperties) {
            if (body.LastChild is Paragraph lastParagraph && lastParagraph.ParagraphProperties?.SectionProperties == null) {
                lastParagraph.ParagraphProperties ??= new ParagraphProperties();
                lastParagraph.ParagraphProperties.SectionProperties = sectionProperties;
                return;
            }

            body.Append(new Paragraph(new ParagraphProperties { SectionProperties = sectionProperties }));
        }

        private List<OpenXmlElement> GetImportedHeaderFooterRoots() {
            var roots = new List<OpenXmlElement>();
            foreach (OpenXmlPart part in _partMap.Values) {
                OpenXmlElement? root = part switch {
                    HeaderPart header => header.Header,
                    FooterPart footer => footer.Footer,
                    _ => null
                };
                if (root != null) {
                    roots.Add(root);
                }
            }

            return roots;
        }

        private static bool IsWordElement(OpenXmlElement element, string localName) {
            return string.Equals(element.LocalName, localName, StringComparison.Ordinal) &&
                   string.Equals(element.NamespaceUri, WordNamespace, StringComparison.Ordinal);
        }

        private static string? GetWordValue(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (string.Equals(attribute.LocalName, "val", StringComparison.Ordinal) &&
                    string.Equals(attribute.NamespaceUri, WordNamespace, StringComparison.Ordinal)) {
                    if (string.IsNullOrEmpty(attribute.Value)) {
                        return null;
                    }

                    return attribute.Value;
                }
            }

            return null;
        }

        private static void SetWordValue(OpenXmlElement element, string value) {
            element.SetAttribute(new OpenXmlAttribute("w", "val", WordNamespace, value));
        }
    }
}
