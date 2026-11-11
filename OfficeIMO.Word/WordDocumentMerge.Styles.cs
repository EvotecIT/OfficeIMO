using System.Xml;
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

        private readonly Dictionary<string, string> _styleMap = new(StringComparer.Ordinal);
        private readonly List<(Style Source, string Id, string? Name, Style? Replaces)> _styleCopies = new();
        private Styles? _sourceStyles;
        private Styles? _destinationStyles;
        private bool _foldSourceDefaults;

        /// <summary>
        /// Maps source styles used by the appended content and their dependencies, copying or renaming where needed.
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

            var usedIds = new HashSet<string>(destinationById.Keys, StringComparer.Ordinal);
            var usedNames = new HashSet<string>(
                _destinationStyles.Elements<Style>().Select(_ => _.StyleName?.Val?.Value).OfType<string>(),
                StringComparer.OrdinalIgnoreCase);

            WordStyleConflictMode mode = _options.StyleConflictMode;
            bool keepSourceLook = mode != WordStyleConflictMode.KeepTarget;
            _foldSourceDefaults = keepSourceLook &&
                !string.Equals(Canonicalize(_sourceStyles.DocDefaults), Canonicalize(_destinationStyles.DocDefaults), StringComparison.Ordinal);
            HashSet<string> usedSourceStyles = CollectUsedSourceStyles(sourceById);

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

            var overwritten = new HashSet<Style>();
            var equivalence = new Dictionary<string, bool>(StringComparer.Ordinal);
            var comparing = new HashSet<string>(StringComparer.Ordinal);

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

                // List ids are package-local and equal ids need not describe equal lists.
                // Keep such styles separate unless list equivalence can be established.
                if (source.Descendants<NumberingId>().Any() || destination.Descendants<NumberingId>().Any()) {
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
                    return string.Equals(sourceBase, destinationBase, StringComparison.Ordinal);
                }

                return ReferenceEquals(FindMatch(sourceParent), destinationParent) && AreEquivalent(sourceParent, destinationParent);
            }

            foreach (Style style in _sourceStyles.Elements<Style>()) {
                string? id = style.StyleId?.Value;
                if (string.IsNullOrEmpty(id) || !usedSourceStyles.Contains(id!) || _styleMap.ContainsKey(id!)) {
                    continue;
                }

                string sourceId = id!;
                Style? match = FindMatch(style);
                if (match != null) {
                    if (!keepSourceLook || AreEquivalent(style, match)) {
                        _styleMap[sourceId] = match.StyleId!.Value!;
                        continue;
                    }

                    if (mode == WordStyleConflictMode.Overwrite) {
                        // Two source styles can match one destination style (one by name, one by id); the first replaces it.
                        _styleMap[sourceId] = match.StyleId!.Value!;
                        if (overwritten.Add(match)) {
                            _styleCopies.Add((style, match.StyleId!.Value!, null, match));
                        }
                        continue;
                    }
                } else if (!IsNameTaken(style, usedNames) && usedIds.Add(sourceId)) {
                    string? name = style.StyleName?.Val?.Value;
                    if (name != null) {
                        usedNames.Add(name);
                    }

                    _styleMap[sourceId] = sourceId;
                    _styleCopies.Add((style, sourceId, null, null));
                    continue;
                }

                (string newId, string newName) = CreateUniqueStyleIdentity(style, usedIds, usedNames);
                _styleMap[sourceId] = newId;
                _styleCopies.Add((style, newId, newName, null));
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

            foreach ((Style source, string id, string? name, Style? replaces) in _styleCopies) {
                var clone = (Style)source.CloneNode(true);
                clone.StyleId = id;
                if (name != null) {
                    clone.StyleName = new StyleName { Val = name };
                    clone.CustomStyle = true;
                }

                StyleValues? type = clone.Type?.Value;
                if (replaces != null) {
                    // The destination keeps its identity: its default choice and its name, which another style may share with the source.
                    clone.Default = replaces.Default?.Value;
                    clone.StyleName = (StyleName?)replaces.StyleName?.CloneNode(true);
                } else if (clone.Default?.Value == true && (type == null || !defaultTypes.Add(type.Value))) {
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
                if (replaces != null) {
                    ReplaceDefinition(replaces, clone);
                } else {
                    _destinationStyles.Append(clone);
                }
            }

            // OfficeIMO reads the saved style part when it checks the style catalog, so flush the changes.
            _destinationStyles.Save();
        }

        /// <summary>
        /// Moves the overwriting definition into the destination's own style element, so <see cref="WordStyle"/>
        /// objects obtained before the merge keep working and read the new definition.
        /// </summary>
        private static void ReplaceDefinition(Style target, Style definition) {
            foreach (KeyValuePair<string, string> declaration in definition.NamespaceDeclarations) {
                if (target.LookupNamespace(declaration.Key) == null) {
                    target.AddNamespaceDeclaration(declaration.Key, declaration.Value);
                }
            }

            target.ClearAllAttributes();
            target.RemoveAllChildren();
            foreach (OpenXmlAttribute attribute in definition.GetAttributes()) {
                target.SetAttribute(attribute);
            }

            foreach (OpenXmlElement child in definition.ChildElements.ToList()) {
                child.Remove();
                target.AppendChild(child);
            }
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
            if (string.Equals(mapped, destinationDefault, StringComparison.Ordinal)) {
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
            var used = new HashSet<string>(StringComparer.Ordinal);

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
            void AddPartReferences(OpenXmlPart? part) {
                if (part == null) return;
                if (part.IsRootElementLoaded) {
                    AddReferences(part.RootElement);
                    return;
                }

                // Keep unsaved live edits above, but do not materialize unrelated cold source DOMs.
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit
                });
                while (reader.Read()) {
                    if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordNamespace &&
                        StyleReferenceElements.Contains(reader.LocalName) &&
                        reader.GetAttribute("val", WordNamespace) is { Length: > 0 } value) {
                        used.Add(value);
                    }
                }
            }

            foreach (HeaderPart part in _sourceMain.HeaderParts) AddPartReferences(part);
            foreach (FooterPart part in _sourceMain.FooterParts) AddPartReferences(part);
            AddPartReferences(_sourceMain.FootnotesPart);
            AddPartReferences(_sourceMain.EndnotesPart);
            AddPartReferences(_sourceMain.WordprocessingCommentsPart);
            AddPartReferences(_sourceMain.NumberingDefinitionsPart);

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

        /// <summary>
        /// Style names are unique across style types, so a source style whose name a destination style of another
        /// type already uses is copied under a new name.
        /// </summary>
        private static bool IsNameTaken(Style style, HashSet<string> usedNames) {
            string? name = style.StyleName?.Val?.Value;
            return name != null && usedNames.Contains(name);
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
            var index = new Dictionary<string, Style>(StringComparer.Ordinal);
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

            // StyleValues is a struct whose ToString() does not name the type, so key on the attribute text.
            return (style.Type?.InnerText ?? string.Empty) + "|" + name!.Trim();
        }

        private static string CanonicalizeStyle(Style style) {
            var builder = new StringBuilder();
            builder.Append(style.Type?.InnerText);
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
