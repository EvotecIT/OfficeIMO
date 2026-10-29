using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal sealed partial class WordDocumentMerge {
        // Style children that describe how a style is presented in the UI rather than how text is formatted.
        private static readonly HashSet<string> PresentationStyleChildren = new(StringComparer.Ordinal) {
            "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden",
            "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid"
        };

        private static readonly HashSet<string> StyleReferenceElements = new(StringComparer.Ordinal) {
            "pStyle", "rStyle", "tblStyle", "styleLink", "numStyleLink"
        };

        private readonly Dictionary<string, string> _styleMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<(Style Source, string Id, string? Name)> _styleCopies = new();
        private Styles? _sourceStyles;
        private Styles? _destinationStyles;
        private bool _foldSourceDefaults;

        /// <summary>
        /// Decides, for every source style, whether it maps onto a destination style or is copied (possibly renamed).
        /// </summary>
        private void PlanStyles() {
            _sourceStyles = _sourceMain.StyleDefinitionsPart?.Styles;
            if (_sourceStyles == null) {
                return;
            }

            _destinationStyles = EnsureDestinationStyles();
            Dictionary<string, Style> sourceById = IndexStylesById(_sourceStyles);
            Dictionary<string, Style> destinationById = IndexStylesById(_destinationStyles);
            var destinationByName = new Dictionary<string, Style>(StringComparer.OrdinalIgnoreCase);
            foreach (Style style in _destinationStyles.Elements<Style>()) {
                string? key = GetStyleNameKey(style);
                if (key != null && !destinationByName.ContainsKey(key)) {
                    destinationByName.Add(key, style);
                }
            }

            var usedIds = new HashSet<string>(destinationById.Keys, StringComparer.OrdinalIgnoreCase);
            var usedNames = new HashSet<string>(
                _destinationStyles.Elements<Style>().Select(_ => _.StyleName?.Val?.Value).OfType<string>(),
                StringComparer.OrdinalIgnoreCase);

            bool keepDifferent = _options.StyleMode == WordMergeStyleMode.KeepDifferentStyles;
            _foldSourceDefaults = keepDifferent &&
                !string.Equals(Canonicalize(_sourceStyles.DocDefaults), Canonicalize(_destinationStyles.DocDefaults), StringComparison.Ordinal);
            HashSet<string> usedSourceStyles = keepDifferent
                ? CollectUsedSourceStyles(sourceById)
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            Style? FindMatch(Style style) {
                string? key = GetStyleNameKey(style);
                if (key != null && destinationByName.TryGetValue(key, out Style? byName)) {
                    return byName;
                }

                string? id = style.StyleId?.Value;
                if (id != null && destinationById.TryGetValue(id, out Style? byId) && byId.Type?.Value == style.Type?.Value) {
                    return byId;
                }

                return null;
            }

            var equivalence = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            var comparing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            bool AreEquivalent(Style source, Style destination) {
                string sourceId = source.StyleId!.Value!;
                if (equivalence.TryGetValue(sourceId, out bool known)) {
                    return known;
                }

                if (!comparing.Add(sourceId)) {
                    return true;
                }

                bool result = CompareStyles(source, destination);
                comparing.Remove(sourceId);
                equivalence[sourceId] = result;
                return result;
            }

            bool CompareStyles(Style source, Style destination) {
                if (source.Type?.Value != destination.Type?.Value) {
                    return false;
                }

                string? sourceBase = source.BasedOn?.Val?.Value;
                string? destinationBase = destination.BasedOn?.Val?.Value;
                if (_foldSourceDefaults && sourceBase == null && source.Type?.Value == StyleValues.Paragraph) {
                    // Root paragraph styles inherit the document defaults, which differ between the documents.
                    return false;
                }

                if (!string.Equals(CanonicalizeStyle(source), CanonicalizeStyle(destination), StringComparison.Ordinal)) {
                    return false;
                }

                if (sourceBase == null || destinationBase == null) {
                    return sourceBase == null && destinationBase == null;
                }

                if (!sourceById.TryGetValue(sourceBase, out Style? sourceParent) ||
                    !destinationById.TryGetValue(destinationBase, out Style? destinationParent)) {
                    return string.Equals(sourceBase, destinationBase, StringComparison.OrdinalIgnoreCase);
                }

                return ReferenceEquals(FindMatch(sourceParent), destinationParent) && AreEquivalent(sourceParent, destinationParent);
            }

            foreach (Style style in _sourceStyles.Elements<Style>()) {
                string? id = style.StyleId?.Value;
                if (string.IsNullOrEmpty(id) || _styleMap.ContainsKey(id!)) {
                    continue;
                }

                string sourceId = id!;
                Style? match = FindMatch(style);
                if (match != null) {
                    if (!keepDifferent || !usedSourceStyles.Contains(sourceId) || AreEquivalent(style, match)) {
                        _styleMap[sourceId] = match.StyleId!.Value!;
                        continue;
                    }
                } else if (usedIds.Add(sourceId)) {
                    string? name = style.StyleName?.Val?.Value;
                    if (name != null) {
                        usedNames.Add(name);
                    }

                    _styleMap[sourceId] = sourceId;
                    _styleCopies.Add((style, sourceId, null));
                    continue;
                }

                (string newId, string newName) = CreateUniqueStyleIdentity(style, usedIds, usedNames);
                _styleMap[sourceId] = newId;
                _styleCopies.Add((style, newId, newName));
            }
        }

        /// <summary>
        /// Copies the planned styles into the destination with their references repointed.
        /// Runs after numbering import so numbering references inside styles can be remapped.
        /// </summary>
        private void CommitStyles() {
            if (_destinationStyles == null) {
                return;
            }

            var defaultTypes = new HashSet<StyleValues>(
                _destinationStyles.Elements<Style>()
                    .Where(_ => _.Default?.Value == true && _.Type?.Value != null)
                    .Select(_ => _.Type!.Value));

            foreach ((Style source, string id, string? name) in _styleCopies) {
                var clone = (Style)source.CloneNode(true);
                clone.StyleId = id;
                if (name != null) {
                    clone.StyleName = new StyleName { Val = name };
                    clone.CustomStyle = true;
                }

                StyleValues? type = clone.Type?.Value;
                if (clone.Default?.Value == true && (type == null || !defaultTypes.Add(type.Value))) {
                    clone.Default = null;
                }

                if (clone.BasedOn?.Val?.Value is string basedOn) {
                    clone.BasedOn.Val = MapStyleId(basedOn);
                }

                if (clone.NextParagraphStyle?.Val?.Value is string next) {
                    clone.NextParagraphStyle.Val = MapStyleId(next);
                }

                if (clone.LinkedStyle?.Val?.Value is string linked) {
                    clone.LinkedStyle.Val = MapStyleId(linked);
                }

                if (_foldSourceDefaults && clone.BasedOn == null && type == StyleValues.Paragraph) {
                    FoldSourceDocumentDefaults(clone);
                }

                RemapNumberingReferences(clone);
                _destinationStyles.Append(clone);
            }

            // OfficeIMO reads the saved style part when it checks the style catalog, so flush the changes.
            _destinationStyles.Save();
        }

        private string MapStyleId(string sourceId) {
            if (_styleMap.TryGetValue(sourceId, out string? mapped)) {
                return mapped;
            }

            return sourceId;
        }

        private void RemapStyleReferences(OpenXmlElement root) {
            if (_styleMap.Count == 0) {
                return;
            }

            foreach (OpenXmlElement element in root.Descendants()) {
                if (!string.Equals(element.NamespaceUri, WordNamespace, StringComparison.Ordinal) ||
                    !StyleReferenceElements.Contains(element.LocalName)) {
                    continue;
                }

                string? value = GetWordValue(element);
                if (value != null && _styleMap.TryGetValue(value, out string? mapped) &&
                    !string.Equals(value, mapped, StringComparison.Ordinal)) {
                    SetWordValue(element, mapped);
                }
            }
        }

        /// <summary>
        /// Paragraphs and tables without an explicit style use the document default style. When the source default
        /// was imported under a different id, the appended content names it explicitly so its formatting is kept.
        /// </summary>
        private void ApplyImplicitDefaultStyles(IEnumerable<OpenXmlElement> roots) {
            if (_sourceStyles == null || _destinationStyles == null) {
                return;
            }

            string? paragraphStyle = GetImportedDefaultStyle(StyleValues.Paragraph);
            string? tableStyle = GetImportedDefaultStyle(StyleValues.Table);
            if (paragraphStyle == null && tableStyle == null) {
                return;
            }

            foreach (OpenXmlElement root in roots) {
                if (paragraphStyle != null) {
                    foreach (Paragraph paragraph in root.Descendants<Paragraph>()) {
                        if (paragraph.ParagraphProperties?.ParagraphStyleId == null) {
                            paragraph.ParagraphProperties ??= new ParagraphProperties();
                            paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = paragraphStyle };
                        }
                    }
                }

                if (tableStyle != null) {
                    foreach (Table table in root.Descendants<Table>()) {
                        TableProperties properties = table.GetFirstChild<TableProperties>() ?? table.PrependChild(new TableProperties());
                        properties.TableStyle ??= new TableStyle { Val = tableStyle };
                    }
                }
            }
        }

        private string? GetImportedDefaultStyle(StyleValues type) {
            string? sourceDefault = GetDefaultStyleId(_sourceStyles!, type);
            if (sourceDefault == null || !_styleMap.TryGetValue(sourceDefault, out string? mapped)) {
                return null;
            }

            string? destinationDefault = GetDefaultStyleId(_destinationStyles!, type);
            if (string.Equals(mapped, destinationDefault, StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            return mapped;
        }

        private static string? GetDefaultStyleId(Styles styles, StyleValues type) {
            return styles.Elements<Style>()
                .FirstOrDefault(_ => _.Type?.Value == type && _.Default?.Value == true)?
                .StyleId?.Value;
        }

        /// <summary>
        /// Copies the source document defaults into a root paragraph style, so text keeps its source formatting
        /// once it is governed by the destination's document defaults.
        /// </summary>
        private void FoldSourceDocumentDefaults(Style style) {
            DocDefaults? defaults = _sourceStyles?.DocDefaults;
            RunPropertiesBaseStyle? runDefaults = defaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
            if (runDefaults != null && runDefaults.HasChildren) {
                style.StyleRunProperties ??= new StyleRunProperties();
                AddMissingChildren(style.StyleRunProperties, runDefaults);
            }

            ParagraphPropertiesBaseStyle? paragraphDefaults = defaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
            if (paragraphDefaults != null && paragraphDefaults.HasChildren) {
                style.StyleParagraphProperties ??= new StyleParagraphProperties();
                AddMissingChildren(style.StyleParagraphProperties, paragraphDefaults);
            }
        }

        private static void AddMissingChildren(OpenXmlCompositeElement target, OpenXmlElement defaults) {
            foreach (OpenXmlElement child in defaults.ChildElements) {
                bool present = target.ChildElements.Any(existing =>
                    string.Equals(existing.LocalName, child.LocalName, StringComparison.Ordinal) &&
                    string.Equals(existing.NamespaceUri, child.NamespaceUri, StringComparison.Ordinal));
                if (!present) {
                    target.AddChild(child.CloneNode(true), throwOnError: false);
                }
            }
        }

        private HashSet<string> CollectUsedSourceStyles(Dictionary<string, Style> sourceById) {
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddReferences(OpenXmlElement? root) {
                if (root == null) {
                    return;
                }

                foreach (OpenXmlElement element in root.Descendants()) {
                    if (string.Equals(element.NamespaceUri, WordNamespace, StringComparison.Ordinal) &&
                        StyleReferenceElements.Contains(element.LocalName) &&
                        GetWordValue(element) is string value) {
                        used.Add(value);
                    }
                }
            }

            AddReferences(_sourceMain.Document);
            foreach (HeaderPart part in _sourceMain.HeaderParts) AddReferences(part.Header);
            foreach (FooterPart part in _sourceMain.FooterParts) AddReferences(part.Footer);
            AddReferences(_sourceMain.FootnotesPart?.Footnotes);
            AddReferences(_sourceMain.EndnotesPart?.Endnotes);
            AddReferences(_sourceMain.WordprocessingCommentsPart?.Comments);
            AddReferences(_sourceMain.NumberingDefinitionsPart?.Numbering);

            foreach (Style style in sourceById.Values) {
                if (style.Default?.Value == true && style.StyleId?.Value is string id) {
                    used.Add(id);
                }
            }

            var pending = new Queue<string>(used);
            while (pending.Count > 0) {
                if (!sourceById.TryGetValue(pending.Dequeue(), out Style? style)) {
                    continue;
                }

                foreach (string? related in new[] { style.BasedOn?.Val?.Value, style.LinkedStyle?.Val?.Value, style.NextParagraphStyle?.Val?.Value }) {
                    if (!string.IsNullOrEmpty(related) && used.Add(related!)) {
                        pending.Enqueue(related!);
                    }
                }
            }

            return used;
        }

        private static (string Id, string Name) CreateUniqueStyleIdentity(Style style, HashSet<string> usedIds, HashSet<string> usedNames) {
            string baseId = style.StyleId!.Value!;
            string baseName = style.StyleName?.Val?.Value ?? baseId;
            for (int suffix = 0; ; suffix++) {
                string id = baseId + "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                string name = baseName + "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (!usedIds.Contains(id) && !usedNames.Contains(name)) {
                    usedIds.Add(id);
                    usedNames.Add(name);
                    return (id, name);
                }
            }
        }

        private Styles EnsureDestinationStyles() {
            StyleDefinitionsPart part = _destinationMain.StyleDefinitionsPart ?? _destinationMain.AddNewPart<StyleDefinitionsPart>();
            part.Styles ??= new Styles();
            return part.Styles;
        }

        private static Dictionary<string, Style> IndexStylesById(Styles styles) {
            var index = new Dictionary<string, Style>(StringComparer.OrdinalIgnoreCase);
            foreach (Style style in styles.Elements<Style>()) {
                string? id = style.StyleId?.Value;
                if (!string.IsNullOrEmpty(id) && !index.ContainsKey(id!)) {
                    index.Add(id!, style);
                }
            }

            return index;
        }

        private static string? GetStyleNameKey(Style style) {
            string? name = style.StyleName?.Val?.Value;
            if (string.IsNullOrWhiteSpace(name)) {
                return null;
            }

            return (style.Type?.Value.ToString() ?? string.Empty) + "|" + name!.Trim();
        }

        private static string CanonicalizeStyle(Style style) {
            var builder = new StringBuilder();
            builder.Append(style.Type?.Value.ToString());
            foreach (OpenXmlElement child in style.ChildElements) {
                if (!PresentationStyleChildren.Contains(child.LocalName)) {
                    AppendCanonical(builder, child, ignoreNumberingIds: true);
                }
            }

            return builder.ToString();
        }

        private static string Canonicalize(OpenXmlElement? element) {
            if (element == null) {
                return string.Empty;
            }

            var builder = new StringBuilder();
            AppendCanonical(builder, element, ignoreNumberingIds: false);
            return builder.ToString();
        }

        /// <summary>
        /// Writes an element in a form that ignores attribute order and namespace prefixes.
        /// Numbering instance ids differ between packages for the same list, so styles compare without them.
        /// </summary>
        private static void AppendCanonical(StringBuilder builder, OpenXmlElement element, bool ignoreNumberingIds) {
            if (ignoreNumberingIds && IsWordElement(element, "numId")) {
                return;
            }

            builder.Append('<').Append(element.NamespaceUri).Append(':').Append(element.LocalName);
            foreach (OpenXmlAttribute attribute in element.GetAttributes()
                         .OrderBy(_ => _.NamespaceUri, StringComparer.Ordinal)
                         .ThenBy(_ => _.LocalName, StringComparer.Ordinal)) {
                builder.Append(' ').Append(attribute.NamespaceUri).Append(':').Append(attribute.LocalName)
                    .Append("=\"").Append(attribute.Value).Append('"');
            }

            builder.Append('>');
            if (element is OpenXmlLeafTextElement text) {
                builder.Append(text.Text);
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendCanonical(builder, child, ignoreNumberingIds);
            }

            builder.Append("</>");
        }
    }
}
