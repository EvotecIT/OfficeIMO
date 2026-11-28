using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    public partial class WordTableOfContent {
        private const string DefaultTocInstruction = " TOC \\o \"1-3\" \\h \\z \\u ";
        private const string DefaultNoContentText = "No table of contents entries found.";

        /// <summary>
        /// Regenerates visible table-of-contents entries from body heading paragraphs.
        /// </summary>
        /// <remarks>
        /// This deterministic refresh uses explicit page breaks to estimate page numbers. Entries belong to the TOC
        /// field result so Word can replace them during a full layout-aware update when the document is opened.
        /// </remarks>
        /// <returns>A report describing the generated entries and skipped headings.</returns>
        public WordTableOfContentRefreshReport RefreshEntries() {
            string instruction = GetTocInstruction();
            string noContentText = GetNoContentText();
            ISet<int> pageNumberSuppressedLevels = GetTocPageNumberSuppressedLevels(instruction);
            string pageNumberSeparator = GetTocPageNumberSeparator(instruction);
            TocHeadingEntry[] included = CollectIncludedEntries(instruction, out int skipped);

            ReplaceContentWithFieldResult(instruction, noContentText, included, pageNumberSuppressedLevels, pageNumberSeparator);
            QueueUpdateOnOpen(force: true);

            return new WordTableOfContentRefreshReport(
                included.Select(heading => heading.ToPublicEntry()).ToArray(),
                skipped,
                "Estimated from explicit page breaks; Word may recalculate layout on open.");
        }

        /// <summary>
        /// Rewrites the table as a complete TOC field whose result holds one hyperlinked entry per heading, with the
        /// page number of each entry in a <c>PAGEREF</c> field, as Word writes it. Page numbers start as explicit-break
        /// estimates; field update refreshes the <c>PAGEREF</c> fields.
        /// </summary>
        internal WordTableOfContentRefreshReport RegenerateFieldResult() {
            string instruction = GetTocInstruction();
            string noContentText = GetNoContentText();
            ISet<int> pageNumberSuppressedLevels = GetTocPageNumberSuppressedLevels(instruction);
            string pageNumberSeparator = GetTocPageNumberSeparator(instruction);
            TocHeadingEntry[] included = CollectIncludedEntries(instruction, out int skipped);

            ReplaceContentWithFieldResult(instruction, noContentText, included, pageNumberSuppressedLevels, pageNumberSeparator);

            return new WordTableOfContentRefreshReport(
                included.Select(_ => _.ToPublicEntry()).ToArray(),
                skipped,
                "Written as a TOC field result; page numbers are PAGEREF fields refreshed by field update.");
        }

        /// <summary>
        /// Finds every table-of-contents content control in the document body, in document order.
        /// </summary>
        internal static IReadOnlyList<WordTableOfContent> GetTablesOfContents(WordDocument document) {
            Body? body = document._wordprocessingDocument.MainDocumentPart?.Document?.Body;
            var tables = new List<WordTableOfContent>();
            if (body == null) {
                return tables;
            }

            foreach (SdtBlock block in body.Descendants<SdtBlock>().ToList()) {
                string? gallery = block.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value;
                if (string.Equals(gallery, "Table of Contents", StringComparison.Ordinal)) {
                    tables.Add(new WordTableOfContent(document, block, queueUpdateOnOpen: false));
                }
            }

            if (tables.Count == 0 && document.TableOfContent is WordTableOfContent imported) {
                tables.Add(imported);
            }

            return tables;
        }

        private TocHeadingEntry[] CollectIncludedEntries(string instruction, out int skipped) {
            if (ContentNode.IsNone) {
                throw new InvalidOperationException("Table of contents content block is missing.");
            }

            string? tcEntryTypeFilter = GetTocEntryTypeFilter(instruction);
            string? bookmarkScope = GetTocBookmarkScope(instruction);
            IReadOnlyDictionary<string, int> customStyleLevels = GetTocCustomStyleLevels(instruction);
            TocSourceOptions sourceOptions = GetTocSourceOptions(instruction, customStyleLevels);
            var headings = CollectHeadingEntries(tcEntryTypeFilter, customStyleLevels, bookmarkScope, sourceOptions).ToArray();
            var included = headings
                .Where(heading => !heading.ApplyLevelRange || (heading.Level >= MinLevel && heading.Level <= MaxLevel))
                .ToArray();
            skipped = headings.Length - included.Length;
            return included;
        }

        private IReadOnlyList<TocHeadingEntry> CollectHeadingEntries(
            string? tcEntryTypeFilter,
            IReadOnlyDictionary<string, int> customStyleLevels,
            string? bookmarkScope,
            TocSourceOptions sourceOptions) {
            const int W = TocNamespace;
            const string Word2010 = "http://schemas.microsoft.com/office/word/2010/wordml";
            OfficeXmlNode body = WordFieldInventory.BodyNode(_document);
            if (body.IsNone) throw new InvalidOperationException("Document body is missing.");
            if (_document._ownedModel is { } model) WordOwnedEdits.PrepareBodyEdit(model);
            HashSet<OfficeXmlNode>? scopedParagraphs = bookmarkScope == null
                ? null : FindScopedHeadingParagraphs(body, bookmarkScope);

            var entries = new List<TocHeadingEntry>();
            var bookmarks = new TocBookmarkState(body);
            var styles = new Dictionary<string, OfficeXmlNode>(StringComparer.Ordinal);
            foreach (OfficeXmlNode style in GetTocStylesRoot().Children(W, "style")) {
                if (style.Attribute(W, "styleId") is { } id && !styles.ContainsKey(id)) styles.Add(id, style);
            }

            OfficeXmlNode control = SdtNode;
            int pageNumber = 1;
            foreach (OfficeXmlNode child in body.Children()) {
                if (child == control) continue;
                if (child.Is(W, "p")) {
                    ProcessParagraph(child);
                } else if (child.Is(W, "tbl") || child.Is(W, "sdt")) {
                    foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(child)) {
                        if (node.Is(W, "p") && !node.AncestorOrSelf(W, "txbxContent").Exists) ProcessParagraph(node);
                    }
                }
            }

            return entries;

            void ProcessParagraph(OfficeXmlNode paragraph) {
                OfficeXmlNode properties = paragraph.Child(W, "pPr");
                OfficeXmlNode pageBreak = properties.Child(W, "pageBreakBefore");
                if (pageBreak.Exists && (pageBreak.Attribute(W, "val") == null || WordFieldInventory.ReadFlag(pageBreak, "val"))) pageNumber++;

                if (!paragraph.AncestorOrSelf(W, "txbxContent").Exists) {
                    HashSet<string>? seenIds = null;
                    foreach (OfficeXmlNode candidate in WordFieldInventory.EnumerateFieldContent(paragraph)) {
                        if (!candidate.Is(W, "p") || !candidate.AncestorOrSelf(W, "txbxContent").Exists ||
                            GetHeadingLevel(candidate, customStyleLevels, sourceOptions, styles) == null) continue;
                        string? id = candidate.Attributes().FirstOrDefault(attr => attr.NamespaceUri == Word2010 && attr.LocalName == "paraId").Value;
                        if (!string.IsNullOrWhiteSpace(id) && !(seenIds ??= new HashSet<string>(StringComparer.Ordinal)).Add(id!))
                            continue;
                        ProcessEntries(candidate);
                    }
                }
                ProcessEntries(paragraph);
                foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(paragraph)) {
                    if (node.Is(W, "br") && node.Attribute(W, "type") == "page") pageNumber++;
                }
                if (StartsNewPage(properties.Child(W, "sectPr"))) pageNumber++;
            }

            void ProcessEntries(OfficeXmlNode paragraph) {
                if (scopedParagraphs != null && !scopedParagraphs.Contains(paragraph)) return;

                if (GetHeadingLevel(paragraph, customStyleLevels, sourceOptions, styles) is { } heading) {
                    string text = GetParagraphText(paragraph);
                    if (!string.IsNullOrWhiteSpace(text)) {
                        RemoveImportedTableTocBookmarks(paragraph, bookmarks);
                        string bookmarkName = EnsureBookmark(paragraph, bookmarks, entries.Count);
                        entries.Add(new TocHeadingEntry(text, heading.Level, pageNumber, bookmarkName, heading.ApplyLevelRange));
                    }
                }

                if (!sourceOptions.IncludeTocEntryFields) return;
                foreach (string instruction in EnumerateTocEntryInstructions(paragraph)) {
                    if (TryCreateTocEntry(instruction, tcEntryTypeFilter, paragraph, bookmarks, entries.Count, pageNumber, out TocHeadingEntry? entry))
                        entries.Add(entry!);
                }
            }
        }

        private static HashSet<OfficeXmlNode> FindScopedHeadingParagraphs(OfficeXmlNode body, string name) {
            const int W = TocNamespace;
            var paragraphs = new HashSet<OfficeXmlNode>();
            OfficeXmlNode start = WordFieldInventory.EnumerateFieldContent(body).FirstOrDefault(node =>
                node.Is(W, "bookmarkStart") && node.Attribute(W, "name") == name);
            string? id = start.Attribute(W, "id");
            if (string.IsNullOrWhiteSpace(id)) return paragraphs;
            OfficeXmlNode end = WordFieldInventory.EnumerateFieldContent(body).FirstOrDefault(node =>
                node.Is(W, "bookmarkEnd") && node.Attribute(W, "id") == id);
            if (end.IsNone) return paragraphs;

            bool inRange = false;
            foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(body)) {
                if (node == start) inRange = true;
                if (node == end) break;
                if (inRange && node.Is(W, "p")) paragraphs.Add(node);
            }

            OfficeXmlNode first = FindBookmarkBoundaryParagraph(start, atStart: true);
            OfficeXmlNode last = FindBookmarkBoundaryParagraph(end, atStart: false);
            if (first.Exists) paragraphs.Add(first);
            if (last.Exists) paragraphs.Add(last);
            return paragraphs;
        }

        private static OfficeXmlNode FindBookmarkBoundaryParagraph(OfficeXmlNode bookmark, bool atStart) {
            const int W = TocNamespace;
            OfficeXmlNode paragraph = bookmark.AncestorOrSelf(W, "p");
            if (paragraph.Exists) return paragraph;
            for (OfficeXmlNode sibling = atStart ? bookmark.NextSibling : bookmark.PreviousSibling; sibling.Exists;
                 sibling = atStart ? sibling.NextSibling : sibling.PreviousSibling) {
                if (sibling.Is(W, "p")) return sibling;
                IEnumerable<OfficeXmlNode> nested = WordFieldInventory.EnumerateFieldContent(sibling)
                    .Where(node => node.Is(W, "p"));
                paragraph = atStart ? nested.FirstOrDefault() : nested.LastOrDefault();
                if (paragraph.Exists) return paragraph;
            }
            return default;
        }

        private static IEnumerable<string> EnumerateTocEntryInstructions(OfficeXmlNode paragraph) {
            const int W = TocNamespace;
            OfficeXmlNode[] content = WordFieldInventory.EnumerateFieldContent(paragraph).ToArray();
            foreach (OfficeXmlNode node in content) {
                if (node.Is(W, "fldSimple") && node.Attribute(W, "instr") is { } instruction && IsTocEntryInstruction(instruction))
                    yield return instruction;
            }

            List<string>? current = null;
            foreach (OfficeXmlNode node in content) {
                if (node.Is(W, "fldChar")) {
                    string? marker = node.Attribute(W, "fldCharType");
                    if (marker == "begin") {
                        current = new List<string>();
                    } else if (marker == "end" && current != null) {
                        string instruction = string.Concat(current);
                        if (IsTocEntryInstruction(instruction)) yield return instruction;
                        current = null;
                    }
                } else if (node.Is(W, "instrText")) {
                    string instruction = node.InnerText;
                    if (current != null) current.Add(instruction);
                    else if (IsTocEntryInstruction(instruction)) yield return instruction;
                }
            }
        }

        private static bool IsTocEntryInstruction(string? instruction) {
            if (string.IsNullOrWhiteSpace(instruction)) {
                return false;
            }

            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction!);
            return parsed.FieldType == WordFieldType.TC;
        }

        private bool TryCreateTocEntry(
            string instruction,
            string? tcEntryTypeFilter,
            OfficeXmlNode paragraph,
            TocBookmarkState bookmarks,
            int entryIndex,
            int pageNumber,
            out TocHeadingEntry? entry) {
            entry = null;
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.TC || parsed.Diagnostics.Count > 0) {
                return false;
            }

            string text = TrimFieldArgument(parsed.Instructions.FirstOrDefault());
            if (text.Length == 0) {
                return false;
            }

            string? entryType = GetSingleSwitchValue(parsed, "\\f", normalizeQuotes: true);
            if (!string.IsNullOrEmpty(tcEntryTypeFilter) &&
                !string.Equals(entryType, tcEntryTypeFilter, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int level = 1;
            string? levelValue = GetSingleSwitchValue(parsed, "\\l", normalizeQuotes: true);
            if (levelValue != null) {
                if (!int.TryParse(levelValue, NumberStyles.None, CultureInfo.InvariantCulture, out int parsedLevel)) {
                    return false;
                }

                level = ClampLevel(parsedLevel);
            }

            string bookmarkName = EnsureBookmark(paragraph, bookmarks, entryIndex);
            entry = new TocHeadingEntry(text, level, pageNumber, bookmarkName, applyLevelRange: false);
            return true;
        }

        private static bool StartsNewPage(SectionProperties? sectionProperties) =>
            StartsNewPage(OfficeXmlNode.FromSdk(sectionProperties));

        private static bool StartsNewPage(OfficeXmlNode section) => section.Exists &&
            section.Child(TocNamespace, "type").Attribute(TocNamespace, "val") is null or "nextPage" or "oddPage" or "evenPage";

        private Paragraph? GetTitleParagraph() => (Paragraph?)GetTitleParagraphNode().SdkElement;

        private OfficeXmlNode GetTitleParagraphNode() {
            // Paragraphs from the table's field onwards belong to its result. A title can contain other fields.
            OfficeXmlNode generatedField = EnumerateFieldInstructions(ContentNode)
                .Where(field => GetFieldType(field.Instruction) is WordFieldType.TOC or WordFieldType.Index)
                .Select(field => field.Start).FirstOrDefault();
            foreach (OfficeXmlNode paragraph in ContentNode.Children(TocNamespace, "p")) {
                if (generatedField.IsInside(paragraph)) return default;
                if (!IsGeneratedResultStyle(paragraph.Child(TocNamespace, "pPr").Child(TocNamespace, "pStyle")
                    .Attribute(TocNamespace, "val"))) return paragraph;
            }
            return default;
        }

        private static bool IsGeneratedResultStyle(string? styleId) {
            if (string.IsNullOrWhiteSpace(styleId)) {
                return false;
            }

            return Regex.IsMatch(styleId!, "^(TOC|Index)[1-9]$") ||
                   string.Equals(styleId, "TableofFigures", StringComparison.Ordinal);
        }

        // Set when this instance wrote heading entries during a save, so later saves keep them current.
        private bool _entriesWrittenOnSave;

        /// <summary>
        /// Writes heading entries on save for a table of contents this library created, while it
        /// still holds only its placeholder or entries an earlier save wrote. Caption lists, indexes
        /// and entries the caller refreshed or edited are left alone.
        /// </summary>
        internal void RefreshEntriesOnSave() {
            const int W = TocNamespace;
            if (ContentNode.IsNone) return;
            var instructions = WordFieldInventory.EnumerateFieldContent(SdtNode)
                .Select(node => node.Is(W, "fldSimple") ? node.Attribute(W, "instr") :
                    node.Is(W, "instrText") ? node.InnerText : null)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .Select(text => text!.Trim())
                .Where(text => !text.StartsWith("PAGEREF", StringComparison.OrdinalIgnoreCase) &&
                    !text.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (instructions.Count == 0 ||
                instructions.Any(text => !text.StartsWith("TOC", StringComparison.OrdinalIgnoreCase) ||
                    Regex.IsMatch(text, @"\\[cfl]\b", RegexOptions.IgnoreCase))) return;
            if (!_entriesWrittenOnSave && WordFieldInventory.EnumerateFieldContent(ContentNode)
                .Any(style => style.Is(W, "pStyle") && style.Attribute(W, "val") is { } id &&
                    (WordBuiltInStyleNames.GetTocLevel(_document._ownedModel is { } model ?
                        model.Styles.GetBuiltInStyleId(id) : _document.GetBuiltInStyleId(id)) > 0 ||
                        id.StartsWith("TOC", StringComparison.Ordinal) && id.Length > 3 && char.IsDigit(id[3])))) return;
            RefreshEntries();
            _entriesWrittenOnSave = true;
        }

        private string GetTocInstruction() {
            const int W = TocNamespace;
            OfficeXmlNode[] markup = WordFieldInventory.EnumerateFieldContent(SdtNode).ToArray();
            foreach (OfficeXmlNode node in markup) {
                if (node.Is(W, "fldSimple") && node.Attribute(W, "instr") is { } instruction && IsTocInstruction(instruction))
                    return instruction;
            }

            foreach (OfficeXmlNode paragraph in markup) {
                if (!paragraph.Is(W, "p")) continue;
                foreach (string instruction in EnumerateComplexFieldInstructions(paragraph)) {
                    if (IsTocInstruction(instruction)) return instruction;
                }
            }

            foreach (OfficeXmlNode node in markup) {
                if (node.Is(W, "instrText") && IsTocInstruction(node.InnerText)) return node.InnerText;
            }

            return DefaultTocInstruction;
        }

        private static bool IsTocInstruction(string? instruction) {
            if (string.IsNullOrWhiteSpace(instruction)) {
                return false;
            }

            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction!);
            return parsed.FieldType == WordFieldType.TOC;
        }

        private static string? GetTocEntryTypeFilter(string instruction) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.TOC || parsed.Diagnostics.Count > 0) {
                return null;
            }

            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\f", StringComparison.OrdinalIgnoreCase))) {
                string value = TrimFieldArgument(fieldSwitch.Trim().Substring(2).Trim());
                return value;
            }

            return null;
        }

        private static string? GetTocBookmarkScope(string instruction) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.TOC || parsed.Diagnostics.Count > 0) {
                return null;
            }

            string? bookmarkName = null;
            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\b", StringComparison.OrdinalIgnoreCase))) {
                if (TryNormalizeIndexBookmarkName(fieldSwitch.Trim().Substring(2).Trim(), out string? parsedBookmarkName)) {
                    bookmarkName = parsedBookmarkName;
                }
            }

            return bookmarkName;
        }

        private IReadOnlyDictionary<string, int> GetTocCustomStyleLevels(string instruction) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.TOC || parsed.Diagnostics.Count > 0) {
                return new Dictionary<string, int>(StringComparer.Ordinal);
            }

            var levels = new Dictionary<string, int>(StringComparer.Ordinal);
            var requestedNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\t", StringComparison.OrdinalIgnoreCase))) {
                string value = TrimFieldArgument(fieldSwitch.Trim().Substring(2).Trim());
                AddCustomStyleLevelPairs(value, levels, requestedNames);
            }

            if (levels.Count == 0) {
                return levels;
            }

            AddCustomStyleAliases(levels, requestedNames);
            return levels;
        }

        private static void AddCustomStyleLevelPairs(string value, Dictionary<string, int> levels, Dictionary<string, int> requestedNames) {
            if (string.IsNullOrWhiteSpace(value)) {
                return;
            }

            string[] parts = value.Split(',');
            for (int index = 0; index + 1 < parts.Length; index += 2) {
                string styleName = TrimFieldArgument(parts[index]);
                string levelValue = parts[index + 1].Trim();
                if (styleName.Length == 0 ||
                    !int.TryParse(levelValue, NumberStyles.None, CultureInfo.InvariantCulture, out int parsedLevel)) {
                    continue;
                }

                int level = ClampLevel(parsedLevel);
                // Keep explicit sources available as exact IDs, including IDs not defined in this document.
                levels[styleName] = level;
                requestedNames[styleName] = level;
            }
        }

        private OfficeXmlNode GetTocStylesRoot() {
            if (_document._ownedModel is { } model) {
                const string relationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "styles";
                string? path = WordOwnedSave.RelatedParts(model, relationshipType).FirstOrDefault();
                if (path == null) return default;
                OfficeXmlTree tree = model.Package.GetXml(path);
                return new OfficeXmlNode(tree, tree.Root);
            }
            return OfficeXmlNode.FromSdk(_document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles);
        }

        private void AddCustomStyleAliases(Dictionary<string, int> levels, IReadOnlyDictionary<string, int> requestedNames) {
            const int W = TocNamespace;
            foreach (OfficeXmlNode style in GetTocStylesRoot().Children(W, "style")) {
                if (style.Attribute(W, "type") != "paragraph") continue;
                string? styleId = style.Attribute(W, "styleId");
                string? name = style.Child(W, "name").Attribute(W, "val");
                if (!string.IsNullOrWhiteSpace(styleId) && !string.IsNullOrWhiteSpace(name) &&
                    requestedNames.TryGetValue(name!, out int level) && !levels.ContainsKey(styleId!)) {
                    levels.Add(styleId!, level);
                }
            }
        }

        private string GetNoContentText(string fallback = DefaultNoContentText) {
            string text = TextNoContent;
            return string.IsNullOrWhiteSpace(text) || (GetPlaceholderPart().IsNone && text == DefaultNoContentText) ? fallback : text;
        }

        private static Paragraph CreateTitleParagraph(string text) {
            return new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        private static Paragraph CreateEntryParagraph(TocHeadingEntry entry, bool includePageNumber = true, string pageNumberSeparator = "\t") {
            var paragraph = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOC" + entry.Level.ToString(CultureInfo.InvariantCulture) }));

            var hyperlink = new Hyperlink {
                Anchor = entry.BookmarkName,
                History = true
            };

            hyperlink.Append(new Run(new Text(entry.Text) { Space = SpaceProcessingModeValues.Preserve }));
            if (includePageNumber) {
                if (pageNumberSeparator == "\t") {
                    hyperlink.Append(new Run(new TabChar()));
                } else if (pageNumberSeparator.Length > 0) {
                    hyperlink.Append(new Run(new Text(pageNumberSeparator) { Space = SpaceProcessingModeValues.Preserve }));
                }

                hyperlink.Append(new Run(new Text(entry.PageNumber.ToString(CultureInfo.InvariantCulture))));
            }

            paragraph.Append(hyperlink);
            return paragraph;
        }

        private static void RemoveIfAttached(OpenXmlElement element) {
            if (element.Parent != null) {
                element.Remove();
            }
        }

        private static (int Level, bool ApplyLevelRange)? GetHeadingLevel(OfficeXmlNode paragraph,
            IReadOnlyDictionary<string, int> customStyleLevels, TocSourceOptions sourceOptions,
            IReadOnlyDictionary<string, OfficeXmlNode> styles) {
            const int W = TocNamespace;
            OfficeXmlNode properties = paragraph.Child(W, "pPr");
            int? outline = OfficeXmlViewValues.ParseInt32(properties.Child(W, "outlineLvl").Attribute(W, "val"));
            if (sourceOptions.IncludeOutlineLevels && outline is >= 0 and <= 8) return (outline.Value + 1, true);

            string? styleId = properties.Child(W, "pStyle").Attribute(W, "val");
            if (string.IsNullOrWhiteSpace(styleId)) return null;
            if (sourceOptions.IncludeCustomStyles && customStyleLevels.TryGetValue(styleId!, out int customLevel))
                return (customLevel, false);

            if (sourceOptions.IncludeHeadingStyles && styleId!.Length == 8 &&
                styleId.StartsWith("Heading", StringComparison.Ordinal) && styleId[7] is >= '1' and <= '9')
                return (styleId[7] - '0', true);
            if (!styles.TryGetValue(styleId!, out OfficeXmlNode style)) return null;

            if (sourceOptions.IncludeHeadingStyles && style.Child(W, "name").Attribute(W, "val") is { } name) {
                Match match = Regex.Match(name, "^heading (?<level>[1-9])(_[0-9]+)?$", RegexOptions.IgnoreCase);
                if (match.Success) return (match.Groups["level"].Value[0] - '0', true);
            }
            if (!sourceOptions.IncludeOutlineLevels) return null;

            var visited = new HashSet<string>(StringComparer.Ordinal);
            while (style.Exists && style.Attribute(W, "styleId") is { } id && visited.Add(id)) {
                outline = OfficeXmlViewValues.ParseInt32(style.Child(W, "pPr").Child(W, "outlineLvl").Attribute(W, "val"));
                if (outline.HasValue) return outline is >= 0 and <= 8 ? (outline.Value + 1, true) : null;
                if (style.Child(W, "basedOn").Attribute(W, "val") is not { } parent || !styles.TryGetValue(parent, out style))
                    break;
            }
            return null;
        }

        private static TocSourceOptions GetTocSourceOptions(string instruction, IReadOnlyDictionary<string, int> customStyleLevels) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.TOC || parsed.Diagnostics.Count > 0) {
                return TocSourceOptions.Default;
            }

            bool hasOutlineRange = HasTocSwitch(parsed, "\\o");
            bool hasOutlineLevels = HasTocSwitch(parsed, "\\u");
            bool hasCustomStyles = HasTocSwitch(parsed, "\\t") && customStyleLevels.Count > 0;
            bool hasTocEntryFields = HasTocSwitch(parsed, "\\f");
            bool hasCaptionSequence = HasTocSwitch(parsed, "\\c");
            bool hasExplicitSource = hasOutlineRange || hasOutlineLevels || hasCustomStyles || hasTocEntryFields || hasCaptionSequence;
            return new TocSourceOptions(
                includeHeadingStyles: hasOutlineRange || !hasExplicitSource,
                includeOutlineLevels: hasOutlineLevels || !hasExplicitSource,
                includeCustomStyles: hasCustomStyles,
                includeTocEntryFields: hasTocEntryFields);
        }

        private static bool HasTocSwitch(WordFieldInventory.ParsedFieldInstruction parsed, string switchName) {
            return parsed.Switches.Any(item => item.Trim().StartsWith(switchName, StringComparison.OrdinalIgnoreCase));
        }

        private static string GetParagraphText(Paragraph paragraph) => GetParagraphText(OfficeXmlNode.FromSdk(paragraph));

        private static string GetParagraphText(OfficeXmlNode paragraph) {
            const int W = TocNamespace;
            bool inTextBox = paragraph.AncestorOrSelf(W, "txbxContent").Exists;
            return string.Concat(WordFieldInventory.EnumerateFieldContent(paragraph)
                .Where(node => node.Is(W, "t") && (inTextBox || !node.AncestorOrSelf(W, "txbxContent").Exists))
                .Select(node => node.InnerText));
        }

        private static string? GetSingleSwitchValue(WordFieldInventory.ParsedFieldInstruction parsed, string switchName, bool normalizeQuotes) {
            string? value = null;
            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith(switchName, StringComparison.OrdinalIgnoreCase))) {
                if (value != null) {
                    return null;
                }

                string rawValue = fieldSwitch.Trim().Substring(switchName.Length).Trim();
                value = normalizeQuotes ? TrimFieldArgument(rawValue) : rawValue;
            }

            return value;
        }

        private static string TrimFieldArgument(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return string.Empty;
            }

            string trimmed = value!.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"') {
                return trimmed.Substring(1, trimmed.Length - 2);
            }

            return trimmed;
        }

        private string EnsureBookmark(OfficeXmlNode paragraph, TocBookmarkState bookmarks, int headingIndex) {
            const int W = TocNamespace;
            bool paragraphInTextBox = paragraph.AncestorOrSelf(W, "txbxContent").Exists;
            foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(paragraph)) {
                if (!node.Is(W, "bookmarkStart") || (!paragraphInTextBox && node.AncestorOrSelf(W, "txbxContent").Exists))
                    continue;
                if (node.Attribute(W, "name") is { } name && !string.IsNullOrWhiteSpace(name)) return name;
            }

            string bookmarkName = bookmarks.AllocateName(headingIndex);
            int nextId = bookmarks.AllocateId();
            _document.ObserveBookmarkId(nextId);
            string bookmarkId = nextId.ToString(CultureInfo.InvariantCulture);
            OfficeXmlNode start = paragraph.CreateElement(W, "bookmarkStart");
            start.SetAttribute(W, "name", bookmarkName);
            start.SetAttribute(W, "id", bookmarkId);
            OfficeXmlNode properties = paragraph.Child(W, "pPr");
            paragraph.InsertBefore(start, properties.Exists ? properties.NextSibling : paragraph.FirstChild);

            AppendTocChild(paragraph, "bookmarkEnd").SetAttribute(W, "id", bookmarkId);
            return bookmarkName;
        }

        private static void RemoveImportedTableTocBookmarks(OfficeXmlNode paragraph, TocBookmarkState bookmarks) {
            const int W = TocNamespace;
            OfficeXmlNode table = paragraph.AncestorOrSelf(W, "tbl");
            if (table.IsNone) return;

            foreach (OfficeXmlNode start in WordFieldInventory.EnumerateFieldContent(paragraph)
                .Where(node => node.Is(W, "bookmarkStart")).ToArray()) {
                string? name = start.Attribute(W, "name");
                string? id = start.Attribute(W, "id");
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(id) ||
                    !name!.StartsWith("_Toc", StringComparison.Ordinal)) continue;

                OfficeXmlNode end = WordFieldInventory.EnumerateFieldContent(table).FirstOrDefault(node =>
                    node.Is(W, "bookmarkEnd") && node.Attribute(W, "id") == id);
                if (end.IsNone) continue;

                start.Remove();
                end.Remove();
                bookmarks.Remove(name);
            }
        }

        private sealed class TocBookmarkState {
            private readonly OfficeXmlNode _body;
            private readonly HashSet<string> _names = new(StringComparer.Ordinal);
            private int? _maxId;

            internal TocBookmarkState(OfficeXmlNode body) {
                _body = body;
                _maxId = Scan(includeNames: true);
            }

            internal string AllocateName(int headingIndex) {
                string prefix = "_OfficeIMO_Toc_" + headingIndex.ToString(CultureInfo.InvariantCulture) + "_";
                int suffix = _names.Count;
                string name;
                do {
                    name = prefix + (suffix++).ToString(CultureInfo.InvariantCulture);
                } while (!_names.Add(name));
                return name;
            }

            internal int AllocateId() {
                int id = (_maxId ?? Scan(includeNames: false)) + 1;
                _maxId = id;
                return id;
            }

            internal void Remove(string name) {
                _names.Remove(name);
                _maxId = null;
            }

            private int Scan(bool includeNames) {
                const int W = TocNamespace;
                int maxId = 0;
                foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(_body, reserveInactiveCompatibilityIds: true)) {
                    bool isStart = node.Is(W, "bookmarkStart");
                    if (!isStart && !node.Is(W, "bookmarkEnd")) continue;
                    if (isStart && includeNames && node.Attribute(W, "name") is { } name && !string.IsNullOrWhiteSpace(name))
                        _names.Add(name);
                    if (int.TryParse(node.Attribute(W, "id"), NumberStyles.None, CultureInfo.InvariantCulture, out int id) && id > maxId)
                        maxId = id;
                }
                return maxId;
            }
        }

        private static ISet<int> GetTocPageNumberSuppressedLevels(string instruction) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.TOC || parsed.Diagnostics.Count > 0) {
                return new HashSet<int>();
            }

            var levels = new HashSet<int>();
            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\n", StringComparison.OrdinalIgnoreCase))) {
                AddPageNumberSuppressedLevels(fieldSwitch.Trim().Substring(2).Trim(), levels);
            }

            return levels;
        }

        private static string GetTocPageNumberSeparator(string instruction) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.TOC || parsed.Diagnostics.Count > 0) {
                return "\t";
            }

            string? separator = null;
            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\p", StringComparison.OrdinalIgnoreCase))) {
                string rawValue = fieldSwitch.Trim().Substring(2).Trim();
                if (rawValue.Length == 0) {
                    continue;
                }

                separator = DecodeTocSeparator(TrimFieldArgument(rawValue));
            }

            return separator ?? "\t";
        }

        private static string DecodeTocSeparator(string value) {
            var builder = new System.Text.StringBuilder(value.Length);
            bool escaped = false;

            foreach (char current in value) {
                if (escaped) {
                    switch (current) {
                        case 't':
                            builder.Append('\t');
                            break;
                        case 'n':
                            builder.Append('\n');
                            break;
                        case 'r':
                            builder.Append('\r');
                            break;
                        default:
                            builder.Append(current);
                            break;
                    }

                    escaped = false;
                    continue;
                }

                if (current == '\\') {
                    escaped = true;
                    continue;
                }

                builder.Append(current);
            }

            if (escaped) {
                builder.Append('\\');
            }

            return builder.ToString();
        }

        private static void AddPageNumberSuppressedLevels(string value, HashSet<int> levels) {
            string rangeText = TrimFieldArgument(value);
            if (rangeText.Length == 0) {
                for (int level = 1; level <= 9; level++) {
                    levels.Add(level);
                }

                return;
            }

            string[] parts = rangeText.Split('-');
            if (parts.Length != 2 ||
                !int.TryParse(parts[0].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int startLevel) ||
                !int.TryParse(parts[1].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int endLevel)) {
                return;
            }

            int first = Math.Max(1, Math.Min(startLevel, endLevel));
            int last = Math.Min(9, Math.Max(startLevel, endLevel));
            for (int level = first; level <= last; level++) {
                levels.Add(level);
            }
        }

        private sealed class TocHeadingEntry {
            internal TocHeadingEntry(string text, int level, int pageNumber, string bookmarkName, bool applyLevelRange) {
                Text = text;
                Level = level;
                PageNumber = pageNumber;
                BookmarkName = bookmarkName;
                ApplyLevelRange = applyLevelRange;
            }

            internal string Text { get; }

            internal int Level { get; }

            internal int PageNumber { get; }

            internal string BookmarkName { get; }

            internal bool ApplyLevelRange { get; }

            internal WordTableOfContentEntry ToPublicEntry() {
                return new WordTableOfContentEntry(Text, Level, PageNumber, BookmarkName);
            }
        }

        private readonly struct TocSourceOptions {
            internal static TocSourceOptions Default { get; } = new TocSourceOptions(
                includeHeadingStyles: true,
                includeOutlineLevels: true,
                includeCustomStyles: false,
                includeTocEntryFields: false);

            internal TocSourceOptions(bool includeHeadingStyles, bool includeOutlineLevels, bool includeCustomStyles, bool includeTocEntryFields) {
                IncludeHeadingStyles = includeHeadingStyles;
                IncludeOutlineLevels = includeOutlineLevels;
                IncludeCustomStyles = includeCustomStyles;
                IncludeTocEntryFields = includeTocEntryFields;
            }

            internal bool IncludeHeadingStyles { get; }

            internal bool IncludeOutlineLevels { get; }

            internal bool IncludeCustomStyles { get; }

            internal bool IncludeTocEntryFields { get; }
        }
    }
}
