using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Word {
    public partial class WordTableOfContent {
        private const string DefaultIndexPageNumberMode = "Estimated from explicit page breaks; Word may recalculate layout on open.";

        /// <summary>
        /// Regenerates visible index entries from <c>XE</c> fields that OfficeIMO can evaluate deterministically.
        /// </summary>
        /// <param name="title">Optional title shown above the generated index.</param>
        /// <returns>A report describing the generated index entries.</returns>
        /// <remarks>
        /// This first-pass refresh supports main terms and nested subentry paths, for example <c>XE "Term"</c>,
        /// <c>XE "Term:Subterm"</c>, and <c>XE "Term:Subterm:Detail"</c>, plus cross-reference text from <c>\t</c> switches,
        /// bounded bookmark page ranges from <c>XE \r</c> switches, optional entry-type filtering from <c>INDEX \f</c> / <c>XE \f</c>,
        /// paragraph-level bookmark range filtering from <c>INDEX \b</c>, bounded Latin letter-range filtering from <c>INDEX \p</c>,
        /// bounded Latin alphabetic heading separators from <c>INDEX \h</c>, and bounded imported column count readback from <c>INDEX \c</c>.
        /// </remarks>
        public WordIndexRefreshReport RefreshIndex(string? title = null) {
            return RefreshIndexCore(title, GetExistingIndexRefreshOptions());
        }

        /// <summary>
        /// Regenerates visible index entries from <c>XE</c> fields matching the requested entry type.
        /// </summary>
        /// <param name="title">Optional title shown above the generated index.</param>
        /// <param name="entryType">Optional Word index entry type from <c>XE \f</c> switches. Pass <c>null</c> to include all entry types.</param>
        /// <returns>A report describing the generated index entries.</returns>
        public WordIndexRefreshReport RefreshIndex(string? title, string? entryType) {
            return RefreshIndexCore(title, new IndexRefreshOptions(NormalizeIndexEntryType(entryType), null, null, IndexHeadingSeparator.Default, IndexSeparators.OfficeImoDefault, null));
        }

        /// <summary>
        /// Regenerates visible index entries from body <c>XE</c> fields matching the requested entry type and bookmark scope.
        /// </summary>
        /// <param name="title">Optional title shown above the generated index.</param>
        /// <param name="entryType">Optional Word index entry type from <c>XE \f</c> switches. Pass <c>null</c> to include all entry types.</param>
        /// <param name="bookmarkName">Optional Word bookmark name from an <c>INDEX \b</c> switch. Pass <c>null</c> to scan the whole body.</param>
        /// <returns>A report describing the generated index entries.</returns>
        /// <remarks>
        /// The bookmark filter and <c>XE \r</c> page ranges use deterministic paragraph-level page estimates.
        /// </remarks>
        public WordIndexRefreshReport RefreshIndex(string? title, string? entryType, string? bookmarkName) {
            return RefreshIndexCore(
                title,
                new IndexRefreshOptions(
                    NormalizeIndexEntryType(entryType),
                    NormalizeIndexBookmarkName(bookmarkName),
                    null,
                    IndexHeadingSeparator.Default,
                    IndexSeparators.OfficeImoDefault,
                    null));
        }

        private WordIndexRefreshReport RefreshIndexCore(string? title, IndexRefreshOptions options) {
            if (ContentNode.IsNone) {
                throw new InvalidOperationException("Table of contents content block is missing.");
            }

            _ = ReadTocUpdateOnOpen();
            var entries = CollectIndexEntries(options, out int skippedEntryCount).ToArray();
            ReplaceIndexContent(string.IsNullOrWhiteSpace(title) ? "Index" : title!.Trim(), entries, options);
            QueueUpdateOnOpen(force: true);

            return new WordIndexRefreshReport(
                entries.Select(entry => entry.ToPublicEntry(options.Separators)).ToArray(),
                skippedEntryCount,
                DefaultIndexPageNumberMode,
                options.ColumnCount);
        }

        private IReadOnlyList<IndexEntry> CollectIndexEntries(IndexRefreshOptions options, out int skippedEntryCount) {
            OfficeXmlNode body = WordFieldInventory.BodyNode(_document);
            if (body.IsNone) throw new InvalidOperationException("Document body is missing.");
            if (_document._ownedModel is { } model) WordOwnedEdits.PrepareBodyEdit(model);

            var candidates = new List<IndexEntryCandidate>();
            BookmarkRange? bookmarkRange = options.BookmarkName == null
                ? null
                : FindBookmarkRange(body, options.BookmarkName);
            HashSet<OfficeXmlNode>? scoped = bookmarkRange is { } range ? CollectScopedParagraphs(body, range) : null;
            OfficeXmlNode control = SdtNode;
            Dictionary<OfficeXmlNode, int> paragraphPages = BuildParagraphPageMap(body, control);
            Dictionary<int, int> footnotePages = BuildNoteAnchorPageMap(body, paragraphPages, "footnoteReference");
            Dictionary<int, int> endnotePages = BuildNoteAnchorPageMap(body, paragraphPages, "endnoteReference");
            skippedEntryCount = 0;

            foreach (WordFieldInventory.NodeFieldRoot root in WordFieldInventory.EnumerateStories(_document)) {
                OfficeXmlNode excludedBlock = root.LocationKind == WordFieldLocationKind.Body ? control : default;
                foreach (OfficeXmlNode paragraph in EnumerateTocSourceParagraphs(root.Root, excludedBlock)) {
                    if (options.BookmarkName == null || scoped?.Contains(paragraph) == true) {
                        foreach (string instruction in EnumerateIndexInstructions(paragraph)) {
                            int pageNumber = GetParagraphPageNumber(paragraph, paragraphPages, footnotePages, endnotePages);
                            if (TryCreateIndexCandidate(instruction, pageNumber, bookmarkName => TryGetBookmarkPageRange(body, paragraphPages, bookmarkName, out IndexPageReference range) ? range : null, out IndexEntryCandidate? candidate)) {
                                if (options.EntryTypeFilter != null &&
                                    !string.Equals(candidate!.EntryType, options.EntryTypeFilter, StringComparison.OrdinalIgnoreCase)) {
                                    continue;
                                }

                                if (options.LetterRange.HasValue && !options.LetterRange.Value.Contains(candidate!.Path[0])) {
                                    continue;
                                }

                                candidates.Add(candidate!);
                            } else {
                                skippedEntryCount++;
                            }
                        }
                    }
                }
            }

            return candidates
                .GroupBy(candidate => candidate.Key, StringComparer.OrdinalIgnoreCase)
                .Select(group => IndexEntry.FromCandidates(group))
                .OrderBy(entry => entry.Path, IndexPathComparer.Instance)
                .ToArray();
        }

        private static IEnumerable<string> EnumerateIndexInstructions(OfficeXmlNode paragraph) {
            foreach (OfficeXmlNode field in EnumerateTocElements(paragraph, "fldSimple")) {
                string instruction = field.Attribute(TocNamespace, "instr") ?? string.Empty;
                if (IsIndexEntryInstruction(instruction)) {
                    yield return instruction;
                }
            }

            List<string> complexInstructions = EnumerateComplexFieldInstructions(paragraph).ToList();
            foreach (string instruction in complexInstructions) {
                if (IsIndexEntryInstruction(instruction)) {
                    yield return instruction;
                }
            }

            if (complexInstructions.Count > 0) {
                yield break;
            }

            foreach (OfficeXmlNode field in EnumerateTocElements(paragraph, "instrText")) {
                string instruction = field.InnerText;
                if (IsIndexEntryInstruction(instruction)) {
                    yield return instruction;
                }
            }
        }

        private static IEnumerable<string> EnumerateComplexFieldInstructions(OfficeXmlNode paragraph) {
            const int W = TocNamespace;
            var instructions = new Stack<StringBuilder?>();
            foreach (OfficeXmlNode run in WordFieldInventory.EnumerateFieldContent(paragraph)) {
                if (!run.Is(W, "r")) continue;
                foreach (OfficeXmlNode child in run.Children()) {
                    if (child.Is(W, "fldChar")) {
                        string? type = child.Attribute(W, "fldCharType");
                        if (type == "begin") {
                            instructions.Push(new StringBuilder());
                        } else if (type is "separate" or "end" && instructions.Count > 0) {
                            StringBuilder? instruction = instructions.Pop();
                            if (instruction != null) yield return instruction.ToString();
                            if (type == "separate") instructions.Push(null);
                        }
                    } else if (instructions.Count > 0 && child.Is(W, "instrText")) {
                        instructions.Peek()?.Append(child.InnerText);
                    }
                }
            }
        }

        private static bool IsIndexEntryInstruction(string? instruction) {
            if (string.IsNullOrWhiteSpace(instruction)) {
                return false;
            }

            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction!);
            return parsed.FieldType == WordFieldType.XE;
        }

        private static Dictionary<OfficeXmlNode, int> BuildParagraphPageMap(OfficeXmlNode body, OfficeXmlNode excludedBlock) {
            const int W = TocNamespace;
            var paragraphPages = new Dictionary<OfficeXmlNode, int>();
            int pageNumber = 1;

            foreach (OfficeXmlNode paragraph in EnumerateTocSourceParagraphs(body, excludedBlock)) {
                OfficeXmlNode properties = paragraph.Child(W, "pPr");
                if (HasTocPageBreakBefore(properties)) pageNumber++;
                paragraphPages[paragraph] = pageNumber;
                pageNumber += EnumerateTocElements(paragraph, "br").Count(node => node.Attribute(W, "type") == "page");
                if (StartsNewPage(properties.Child(W, "sectPr"))) pageNumber++;
            }

            return paragraphPages;
        }

        private static Dictionary<int, int> BuildNoteAnchorPageMap(OfficeXmlNode body,
            IReadOnlyDictionary<OfficeXmlNode, int> paragraphPages, string referenceName) {
            var pages = new Dictionary<int, int>();
            foreach (OfficeXmlNode paragraph in EnumerateTocElements(body, "p")) {
                int page = paragraphPages.TryGetValue(paragraph, out int mapped) ? mapped : 1;
                foreach (OfficeXmlNode reference in EnumerateTocElements(paragraph, referenceName)) {
                    if (OfficeXmlViewValues.ParseInt32(reference.Attribute(TocNamespace, "id")) is int id) pages[id] = page;
                }
            }
            return pages;
        }

        private static int GetParagraphPageNumber(OfficeXmlNode paragraph,
            IReadOnlyDictionary<OfficeXmlNode, int> paragraphPages, IReadOnlyDictionary<int, int> footnotePages,
            IReadOnlyDictionary<int, int> endnotePages) {
            if (paragraphPages.TryGetValue(paragraph, out int page)) return page;
            if (OfficeXmlViewValues.ParseInt32(paragraph.AncestorOrSelf(TocNamespace, "footnote").Attribute(TocNamespace, "id")) is int footnote &&
                footnotePages.TryGetValue(footnote, out page)) return page;
            if (OfficeXmlViewValues.ParseInt32(paragraph.AncestorOrSelf(TocNamespace, "endnote").Attribute(TocNamespace, "id")) is int endnote &&
                endnotePages.TryGetValue(endnote, out page)) return page;
            return 1;
        }

        private static IEnumerable<OfficeXmlNode> EnumerateTocSourceParagraphs(OfficeXmlNode container, OfficeXmlNode excludedBlock) {
            const int W = TocNamespace;
            foreach (OfficeXmlNode child in container.Children()) {
                if (child == excludedBlock) continue;
                if (child.Is(W, "p")) {
                    yield return child;
                    continue;
                }
                if (!(child.Is(W, "tbl") || child.Is(W, "sdt") ||
                      (child.Is(W, "footnote") || child.Is(W, "endnote")) && WordNoteResolver.IsOrdinary(child))) continue;
                foreach (OfficeXmlNode paragraph in EnumerateTocElements(child, "p")) {
                    if (paragraph.AncestorOrSelf(W, "txbxContent").IsNone) yield return paragraph;
                }
            }
        }

        private static IEnumerable<OfficeXmlNode> EnumerateTocElements(OfficeXmlNode root, string localName) =>
            WordFieldInventory.EnumerateFieldContent(root).Where(node => node.Is(TocNamespace, localName));

        private static bool HasTocPageBreakBefore(OfficeXmlNode properties) {
            OfficeXmlNode pageBreak = properties.Child(TocNamespace, "pageBreakBefore");
            return pageBreak.Exists && (pageBreak.Attribute(TocNamespace, "val") == null || WordFieldInventory.ReadFlag(pageBreak, "val"));
        }

        private static bool TryCreateIndexCandidate(string instruction, int pageNumber, Func<string, IndexPageReference?> resolvePageRange, out IndexEntryCandidate? candidate) {
            candidate = null;
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.XE || parsed.Diagnostics.Count > 0) {
                return false;
            }

            string? rawTerm = parsed.Instructions.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(rawTerm)) {
                return false;
            }

            string? pageRangeBookmarkName = null;
            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\r", StringComparison.OrdinalIgnoreCase))) {
                if (pageRangeBookmarkName != null) {
                    return false;
                }

                if (!TryNormalizeIndexBookmarkName(fieldSwitch.Trim().Substring(2).Trim(), out pageRangeBookmarkName)) {
                    return false;
                }
            }

            string? crossReferenceText = null;
            string? entryType = null;
            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\t", StringComparison.OrdinalIgnoreCase))) {
                if (crossReferenceText != null) {
                    return false;
                }

                crossReferenceText = TrimIndexQuotes(fieldSwitch.Trim().Substring(2).Trim());
                if (crossReferenceText.Length == 0) {
                    return false;
                }
            }

            if (crossReferenceText != null && pageRangeBookmarkName != null) {
                return false;
            }

            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\f", StringComparison.OrdinalIgnoreCase))) {
                if (entryType != null) {
                    return false;
                }

                if (!TryNormalizeIndexEntryType(fieldSwitch.Trim().Substring(2).Trim(), out entryType)) {
                    return false;
                }
            }

            string[] path = TrimIndexQuotes(rawTerm)
                .Split(':')
                .Select(part => part.Trim())
                .ToArray();
            if (path.Length == 0 || path.Any(part => part.Length == 0)) {
                return false;
            }

            IndexPageReference? pageReference = null;
            if (crossReferenceText == null) {
                pageReference = pageRangeBookmarkName == null
                    ? new IndexPageReference(pageNumber, pageNumber)
                    : resolvePageRange(pageRangeBookmarkName);
                if (pageReference == null) {
                    return false;
                }
            }

            candidate = new IndexEntryCandidate(path, pageReference, crossReferenceText, entryType);
            return true;
        }

        private void ReplaceIndexContent(string title, IReadOnlyList<IndexEntry> entries, IndexRefreshOptions options) {
            string noContentText = GetNoContentText("No index entries found.");
            OfficeXmlNode content = ResetTocContentWithTitle(title);
            var result = new List<OfficeXmlNode>();
            string? previousHeadingKey = null;
            foreach (IndexNode node in BuildIndexTree(entries)) {
                if (options.HeadingSeparator != null &&
                    options.HeadingSeparator.Value.TryCreateHeading(node.Name, out string headingText, out string headingKey) &&
                    !string.Equals(previousHeadingKey, headingKey, StringComparison.Ordinal)) {
                    result.Add(CreateStyledResultParagraph(content, "IndexHeading", headingText));
                    previousHeadingKey = headingKey;
                }
                AppendIndexNode(content, result, node, options.Separators);
            }
            AppendGeneratedFieldResult(content, CreateIndexInstruction(options), result, noContentText);
        }

        private static void AppendIndexNode(OfficeXmlNode content, List<OfficeXmlNode> result, IndexNode root, IndexSeparators separators) {
            var pending = new Stack<IndexNode>();
            pending.Push(root);
            while (pending.Count > 0) {
                IndexNode node = pending.Pop();
                string style = GetIndexStyleId(node.Level);
                if (node.Entries.Count == 0) {
                    result.Add(CreateStyledResultParagraph(content, style, node.Name));
                } else {
                    foreach (IndexEntry entry in node.Entries.OrderBy(item => item.IsCrossReference)) {
                        OfficeXmlNode paragraph = CreateStyledResultParagraph(content, style, node.Name);
                        AppendIndexEntry(paragraph, entry, separators);
                        result.Add(paragraph);
                    }
                }
                foreach (IndexNode child in node.Children.Values.Reverse()) pending.Push(child);
            }
        }

        private static IReadOnlyList<IndexNode> BuildIndexTree(IReadOnlyList<IndexEntry> entries) {
            var roots = new SortedDictionary<string, IndexNode>(StringComparer.OrdinalIgnoreCase);

            foreach (IndexEntry entry in entries) {
                SortedDictionary<string, IndexNode> siblings = roots;
                IndexNode? current = null;

                for (int i = 0; i < entry.Path.Count; i++) {
                    string part = entry.Path[i];
                    if (!siblings.TryGetValue(part, out IndexNode? node)) {
                        node = new IndexNode(part, i + 1);
                        siblings.Add(part, node);
                    }

                    current = node;
                    siblings = node.Children;
                }

                current?.Entries.Add(entry);
            }

            return roots.Values.ToArray();
        }

        private static void AppendIndexEntry(OfficeXmlNode paragraph, IndexEntry entry, IndexSeparators separators) {
            if (entry.IsCrossReference) {
                AppendIndexSeparatorRuns(paragraph, separators.CrossReferenceSeparator);
                paragraph.Append(CreateTocTextRun(paragraph, "t", entry.CrossReferenceText!, preserveSpace: false));
            } else {
                string pageText = string.Join(
                    separators.PageReferenceSeparator,
                    entry.PageReferences.Select(reference => reference.ToDisplayText(separators.PageRangeSeparator)));
                AppendIndexSeparatorRuns(paragraph, separators.EntryPageSeparator);
                paragraph.Append(CreateTocTextRun(paragraph, "t", pageText, preserveSpace: true));
            }
        }

        private static void AppendIndexSeparatorRuns(OfficeXmlNode paragraph, string separator) {
            int textStart = 0;
            for (int i = 0; i < separator.Length; i++) {
                if (separator[i] != '\t') continue;
                if (i > textStart) {
                    paragraph.Append(CreateTocTextRun(paragraph, "t", separator.Substring(textStart, i - textStart), preserveSpace: true));
                }
                AppendTocChild(AppendTocChild(paragraph, "r"), "tab");
                textStart = i + 1;
            }
            if (textStart < separator.Length) {
                paragraph.Append(CreateTocTextRun(paragraph, "t", separator.Substring(textStart), preserveSpace: true));
            }
        }

        private static string GetIndexStyleId(int level) {
            int normalizedLevel = Math.Max(1, Math.Min(9, level));
            return "Index" + normalizedLevel.ToString(CultureInfo.InvariantCulture);
        }

        private IndexRefreshOptions GetExistingIndexRefreshOptions() {
            foreach (OfficeXmlNode simpleField in EnumerateTocElements(SdtNode, "fldSimple")) {
                string? instruction = simpleField.Attribute(TocNamespace, "instr");
                IndexRefreshOptions? options = GetIndexRefreshOptions(instruction ?? string.Empty);
                if (options != null) {
                    return options.Value;
                }
            }

            foreach (OfficeXmlNode paragraph in EnumerateTocElements(SdtNode, "p")) {
                foreach (string instruction in EnumerateComplexFieldInstructions(paragraph)) {
                    IndexRefreshOptions? options = GetIndexRefreshOptions(instruction);
                    if (options != null) {
                        return options.Value;
                    }
                }
            }

            foreach (OfficeXmlNode fieldCode in EnumerateTocElements(SdtNode, "instrText")) {
                IndexRefreshOptions? options = GetIndexRefreshOptions(fieldCode.InnerText);
                if (options != null) {
                    return options.Value;
                }
            }

            return IndexRefreshOptions.Empty;
        }

        private static IndexRefreshOptions? GetIndexRefreshOptions(string instruction) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType != WordFieldType.Index || parsed.Diagnostics.Count > 0) {
                return null;
            }

            string? entryType = null;
            string? bookmarkName = null;
            IndexLetterRange? letterRange = null;
            IndexHeadingSeparator? headingSeparator = null;
            IndexSeparators separators = IndexSeparators.WordDefault;
            int? columnCount = null;
            foreach (string fieldSwitchValue in GetIndexSwitchValues(parsed.Switches, 'f')) {
                if (TryNormalizeIndexEntryType(fieldSwitchValue, out string? parsedEntryType)) {
                    entryType = parsedEntryType;
                }
            }

            foreach (string fieldSwitchValue in GetIndexSwitchValues(parsed.Switches, 'b')) {
                if (TryNormalizeIndexBookmarkName(fieldSwitchValue, out string? parsedBookmarkName)) {
                    bookmarkName = parsedBookmarkName;
                }
            }

            foreach (string fieldSwitchValue in GetIndexSwitchValues(parsed.Switches, 'p')) {
                if (TryParseIndexLetterRange(fieldSwitchValue, out IndexLetterRange parsedLetterRange)) {
                    letterRange = parsedLetterRange;
                }
            }

            foreach (string fieldSwitchValue in GetIndexSwitchValues(parsed.Switches, 'h')) {
                if (TryParseIndexHeadingSeparator(fieldSwitchValue, out IndexHeadingSeparator parsedHeadingSeparator)) {
                    headingSeparator = parsedHeadingSeparator;
                }
            }

            foreach (string fieldSwitchValue in GetIndexSwitchValues(parsed.Switches, 'c')) {
                if (TryParseIndexColumnCount(fieldSwitchValue, out int parsedColumnCount)) {
                    columnCount = parsedColumnCount;
                }
            }

            foreach (string fieldSwitchValue in GetIndexSwitchValues(parsed.Switches, 'e')) {
                if (TryNormalizeIndexSeparator(fieldSwitchValue, out string? entryPageSeparator)) {
                    separators = separators.WithEntryPageSeparator(entryPageSeparator!);
                }
            }

            foreach (string fieldSwitchValue in GetIndexSwitchValues(parsed.Switches, 'l')) {
                if (TryNormalizeIndexSeparator(fieldSwitchValue, out string? pageReferenceSeparator)) {
                    separators = separators.WithPageReferenceSeparator(pageReferenceSeparator!);
                }
            }

            foreach (string fieldSwitchValue in GetIndexSwitchValues(parsed.Switches, 'g')) {
                if (TryNormalizeIndexSeparator(fieldSwitchValue, out string? pageRangeSeparator)) {
                    separators = separators.WithPageRangeSeparator(pageRangeSeparator!);
                }
            }

            foreach (string fieldSwitchValue in GetIndexSwitchValues(parsed.Switches, 'k')) {
                if (TryNormalizeIndexSeparator(fieldSwitchValue, out string? crossReferenceSeparator)) {
                    separators = separators.WithCrossReferenceSeparator(crossReferenceSeparator!);
                }
            }

            return new IndexRefreshOptions(entryType, bookmarkName, letterRange, headingSeparator, separators, columnCount);
        }

        private static string CreateIndexInstruction(IndexRefreshOptions options) {
            string instruction =
                " INDEX \\e \"" + EscapeIndexSwitchValue(options.Separators.EntryPageSeparator) + "\"" +
                " \\l \"" + EscapeIndexSwitchValue(options.Separators.PageReferenceSeparator) + "\"" +
                " \\g \"" + EscapeIndexSwitchValue(options.Separators.PageRangeSeparator) + "\"" +
                " \\k \"" + EscapeIndexSwitchValue(options.Separators.CrossReferenceSeparator) + "\" ";
            if (options.HeadingSeparator != null) {
                instruction += "\\h \"" + EscapeIndexSwitchValue(options.HeadingSeparator.Value.SwitchValue) + "\" ";
            }

            if (options.EntryTypeFilter != null) {
                instruction += "\\f \"" + options.EntryTypeFilter + "\" ";
            }

            if (options.BookmarkName != null) {
                instruction += "\\b \"" + options.BookmarkName + "\" ";
            }

            if (options.LetterRange.HasValue) {
                instruction += "\\p \"" + options.LetterRange.Value.SwitchValue + "\" ";
            }

            if (options.ColumnCount.HasValue) {
                instruction += "\\c \"" + options.ColumnCount.Value.ToString(CultureInfo.InvariantCulture) + "\" ";
            }

            return instruction;
        }

        private static IEnumerable<string> GetIndexSwitchValues(IEnumerable<string> switches, char switchName) {
            foreach (string fieldSwitch in switches) {
                string trimmed = fieldSwitch.Trim();
                if (trimmed.Length < 2 ||
                    trimmed[0] != '\\' ||
                    char.ToUpperInvariant(trimmed[1]) != char.ToUpperInvariant(switchName) ||
                    (trimmed.Length > 2 && !char.IsWhiteSpace(trimmed[2]))) {
                    continue;
                }

                yield return trimmed.Substring(2).Trim();
            }
        }

        private static string? NormalizeIndexEntryType(string? entryType) {
            if (entryType == null) {
                return null;
            }

            if (TryNormalizeIndexEntryType(entryType, out string? normalized)) {
                return normalized;
            }

            throw new ArgumentException("Index entry type must be a single Word XE entry type without whitespace, quotes, or field switches.", nameof(entryType));
        }

        private static bool TryNormalizeIndexEntryType(string value, out string? entryType) {
            entryType = null;
            string normalized = TrimIndexQuotes(value);
            if (normalized.Length == 0 ||
                normalized.Any(char.IsWhiteSpace) ||
                normalized.IndexOf('\\') >= 0 ||
                normalized.IndexOf('"') >= 0) {
                return false;
            }

            entryType = normalized;
            return true;
        }

        private static string? NormalizeIndexBookmarkName(string? bookmarkName) {
            if (bookmarkName == null) {
                return null;
            }

            if (TryNormalizeIndexBookmarkName(bookmarkName, out string? normalized)) {
                return normalized;
            }

            throw new ArgumentException("Index bookmark scope must be a single Word bookmark name without whitespace, quotes, or field switches.", nameof(bookmarkName));
        }

        private static bool TryNormalizeIndexBookmarkName(string value, out string? bookmarkName) {
            bookmarkName = null;
            string normalized = TrimIndexQuotes(value);
            if (normalized.Length == 0 ||
                normalized.Any(char.IsWhiteSpace) ||
                normalized.IndexOf('\\') >= 0 ||
                normalized.IndexOf('"') >= 0) {
                return false;
            }

            bookmarkName = normalized;
            return true;
        }

        private static bool TryNormalizeIndexSeparator(string value, out string? separator) {
            separator = null;
            string normalized = DecodeIndexSeparator(TrimIndexQuotes(value));
            if (normalized.IndexOf('"') >= 0 || normalized.Length > 5) {
                return false;
            }

            separator = normalized;
            return true;
        }

        private static bool TryParseIndexLetterRange(string value, out IndexLetterRange range) {
            range = default;
            string normalized = TrimIndexQuotes(value).Trim().ToUpperInvariant();
            if (normalized.Length == 3 &&
                IsAsciiLetter(normalized[0]) &&
                normalized[1] == '-' &&
                IsAsciiLetter(normalized[2]) &&
                normalized[0] <= normalized[2]) {
                range = new IndexLetterRange(normalized, includeSymbols: false, normalized[0], normalized[2]);
                return true;
            }

            if (normalized.Length == 4 &&
                normalized.StartsWith("!--", StringComparison.Ordinal) &&
                IsAsciiLetter(normalized[3])) {
                range = new IndexLetterRange(normalized, includeSymbols: true, 'A', normalized[3]);
                return true;
            }

            return false;
        }

        private static bool TryParseIndexHeadingSeparator(string value, out IndexHeadingSeparator separator) {
            separator = default;
            string normalized = DecodeIndexSeparator(TrimIndexQuotes(value));
            if (normalized.IndexOf('"') >= 0 || normalized.Length > 40) {
                return false;
            }

            separator = new IndexHeadingSeparator(normalized);
            return true;
        }

        private static bool TryParseIndexColumnCount(string value, out int columnCount) {
            columnCount = 0;
            string normalized = TrimIndexQuotes(value).Trim();
            if (!int.TryParse(normalized, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ||
                parsed < 1 ||
                parsed > 9) {
                return false;
            }

            columnCount = parsed;
            return true;
        }

        private static bool IsAsciiLetter(char value) {
            return value >= 'A' && value <= 'Z';
        }

        private static string DecodeIndexSeparator(string value) {
            return value.Replace("\\t", "\t").Replace("\\T", "\t");
        }

        private static string EscapeIndexSwitchValue(string value) {
            return value.Replace("\t", "\\t");
        }

        private static BookmarkRange? FindBookmarkRange(OfficeXmlNode body, string bookmarkName) {
            const int W = TocNamespace;
            OfficeXmlNode start = EnumerateTocElements(body, "bookmarkStart")
                .FirstOrDefault(node => node.Attribute(W, "name") == bookmarkName);
            string? id = start.Attribute(W, "id");
            if (string.IsNullOrWhiteSpace(id)) return null;

            OfficeXmlNode end = EnumerateTocElements(body, "bookmarkEnd")
                .FirstOrDefault(node => node.Attribute(W, "id") == id);
            return end.IsNone ? null : new BookmarkRange(start, end);
        }

        private static bool TryGetBookmarkPageRange(OfficeXmlNode body, IReadOnlyDictionary<OfficeXmlNode, int> paragraphPages,
            string bookmarkName, out IndexPageReference pageRange) {
            pageRange = default;
            if (FindBookmarkRange(body, bookmarkName) is not { } range) return false;

            OfficeXmlNode start = FindBookmarkBoundaryParagraph(range.Start, atStart: true);
            OfficeXmlNode end = FindBookmarkBoundaryParagraph(range.End, atStart: false);
            if (!paragraphPages.TryGetValue(start, out int startPage) ||
                !paragraphPages.TryGetValue(end, out int endPage)) return false;

            pageRange = new IndexPageReference(startPage, endPage);
            return true;
        }

        private static string TrimIndexQuotes(string value) {
            string trimmed = value.Trim();
            return trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"'
                ? trimmed.Substring(1, trimmed.Length - 2)
                : trimmed;
        }

        private readonly struct IndexRefreshOptions {
            internal static readonly IndexRefreshOptions Empty = new IndexRefreshOptions(null, null, null, IndexHeadingSeparator.Default, IndexSeparators.OfficeImoDefault, null);

            internal IndexRefreshOptions(string? entryTypeFilter, string? bookmarkName, IndexLetterRange? letterRange, IndexHeadingSeparator? headingSeparator, IndexSeparators separators, int? columnCount) {
                EntryTypeFilter = entryTypeFilter;
                BookmarkName = bookmarkName;
                LetterRange = letterRange;
                HeadingSeparator = headingSeparator;
                Separators = separators;
                ColumnCount = columnCount;
            }

            internal string? EntryTypeFilter { get; }

            internal string? BookmarkName { get; }

            internal IndexLetterRange? LetterRange { get; }

            internal IndexHeadingSeparator? HeadingSeparator { get; }

            internal IndexSeparators Separators { get; }

            internal int? ColumnCount { get; }
        }

        private readonly struct IndexLetterRange {
            internal IndexLetterRange(string switchValue, bool includeSymbols, char startLetter, char endLetter) {
                SwitchValue = switchValue;
                IncludeSymbols = includeSymbols;
                StartLetter = startLetter;
                EndLetter = endLetter;
            }

            internal string SwitchValue { get; }

            private bool IncludeSymbols { get; }

            private char StartLetter { get; }

            private char EndLetter { get; }

            internal bool Contains(string term) {
                string trimmed = term.TrimStart();
                if (trimmed.Length == 0) {
                    return false;
                }

                char normalized = char.ToUpperInvariant(trimmed[0]);
                if (!IsAsciiLetter(normalized)) {
                    return IncludeSymbols;
                }

                return normalized >= StartLetter && normalized <= EndLetter;
            }
        }

        private readonly struct IndexHeadingSeparator {
            internal static readonly IndexHeadingSeparator Default = new IndexHeadingSeparator("A");

            internal IndexHeadingSeparator(string switchValue) {
                SwitchValue = switchValue;
            }

            internal string SwitchValue { get; }

            internal bool TryCreateHeading(string term, out string headingText, out string headingKey) {
                headingText = string.Empty;
                headingKey = string.Empty;

                string trimmed = term.TrimStart();
                if (trimmed.Length == 0) {
                    return false;
                }

                char normalized = char.ToUpperInvariant(trimmed[0]);
                if (!IsAsciiLetter(normalized)) {
                    return false;
                }

                headingKey = normalized.ToString(CultureInfo.InvariantCulture);
                headingText = FormatHeadingText(normalized);
                return true;
            }

            private string FormatHeadingText(char normalizedLetter) {
                if (SwitchValue.Length == 0) {
                    return string.Empty;
                }

                return SwitchValue
                    .Replace("A", normalizedLetter.ToString(CultureInfo.InvariantCulture))
                    .Replace("a", char.ToLowerInvariant(normalizedLetter).ToString(CultureInfo.InvariantCulture));
            }
        }

        private readonly struct IndexSeparators {
            internal static readonly IndexSeparators OfficeImoDefault = new IndexSeparators("\t", ", ", "-", ". ");
            internal static readonly IndexSeparators WordDefault = new IndexSeparators(", ", ", ", "-", ". ");

            internal IndexSeparators(string entryPageSeparator, string pageReferenceSeparator, string pageRangeSeparator, string crossReferenceSeparator) {
                EntryPageSeparator = entryPageSeparator;
                PageReferenceSeparator = pageReferenceSeparator;
                PageRangeSeparator = pageRangeSeparator;
                CrossReferenceSeparator = crossReferenceSeparator;
            }

            internal string EntryPageSeparator { get; }

            internal string PageReferenceSeparator { get; }

            internal string PageRangeSeparator { get; }

            internal string CrossReferenceSeparator { get; }

            internal IndexSeparators WithEntryPageSeparator(string value) {
                return new IndexSeparators(value, PageReferenceSeparator, PageRangeSeparator, CrossReferenceSeparator);
            }

            internal IndexSeparators WithPageReferenceSeparator(string value) {
                return new IndexSeparators(EntryPageSeparator, value, PageRangeSeparator, CrossReferenceSeparator);
            }

            internal IndexSeparators WithPageRangeSeparator(string value) {
                return new IndexSeparators(EntryPageSeparator, PageReferenceSeparator, value, CrossReferenceSeparator);
            }

            internal IndexSeparators WithCrossReferenceSeparator(string value) {
                return new IndexSeparators(EntryPageSeparator, PageReferenceSeparator, PageRangeSeparator, value);
            }
        }

        private readonly struct BookmarkRange {
            internal BookmarkRange(OfficeXmlNode start, OfficeXmlNode end) {
                Start = start;
                End = end;
            }

            internal OfficeXmlNode Start { get; }

            internal OfficeXmlNode End { get; }
        }

        private readonly struct IndexPageReference : IEquatable<IndexPageReference>, IComparable<IndexPageReference> {
            internal IndexPageReference(int startPage, int endPage) {
                StartPage = Math.Min(startPage, endPage);
                EndPage = Math.Max(startPage, endPage);
            }

            internal int StartPage { get; }

            internal int EndPage { get; }

            internal bool IsRange => StartPage != EndPage;

            internal string DisplayText => ToDisplayText("-");

            internal string ToDisplayText(string pageRangeSeparator) {
                return IsRange
                ? StartPage.ToString(CultureInfo.InvariantCulture) + pageRangeSeparator + EndPage.ToString(CultureInfo.InvariantCulture)
                : StartPage.ToString(CultureInfo.InvariantCulture);
            }

            public int CompareTo(IndexPageReference other) {
                int startComparison = StartPage.CompareTo(other.StartPage);
                return startComparison != 0 ? startComparison : EndPage.CompareTo(other.EndPage);
            }

            public bool Equals(IndexPageReference other) {
                return StartPage == other.StartPage && EndPage == other.EndPage;
            }

            public override bool Equals(object? obj) {
                return obj is IndexPageReference other && Equals(other);
            }

            public override int GetHashCode() {
                unchecked {
                    return (StartPage * 397) ^ EndPage;
                }
            }
        }

        private sealed class IndexEntryCandidate {
            internal IndexEntryCandidate(IReadOnlyList<string> path, IndexPageReference? pageReference, string? crossReferenceText, string? entryType) {
                Path = path.ToArray();
                PageReference = pageReference;
                CrossReferenceText = crossReferenceText;
                EntryType = entryType;
            }

            internal IReadOnlyList<string> Path { get; }

            internal IndexPageReference? PageReference { get; }

            internal string? CrossReferenceText { get; }

            internal string? EntryType { get; }

            internal string Key => string.Join("\u0000", Path) + "\u0001" + (CrossReferenceText ?? string.Empty) + "\u0001" + (EntryType ?? string.Empty);
        }

        private sealed class IndexEntry {
            private IndexEntry(IReadOnlyList<string> path, IReadOnlyList<IndexPageReference> pageReferences, string? crossReferenceText, string? entryType) {
                Path = path.ToArray();
                PageReferences = pageReferences.ToArray();
                CrossReferenceText = crossReferenceText;
                EntryType = entryType;
            }

            internal string Term => Path[0];

            internal string? Subterm => Path.Count > 1 ? Path[1] : null;

            internal IReadOnlyList<string> Path { get; }

            internal IReadOnlyList<string> Subterms => Path.Skip(1).ToArray();

            internal IReadOnlyList<IndexPageReference> PageReferences { get; }

            internal IReadOnlyList<int> PageNumbers => PageReferences
                .Where(reference => !reference.IsRange)
                .Select(reference => reference.StartPage)
                .ToArray();

            internal string? CrossReferenceText { get; }

            internal string? EntryType { get; }

            internal bool IsCrossReference => !string.IsNullOrWhiteSpace(CrossReferenceText);

            internal static IndexEntry FromCandidates(IEnumerable<IndexEntryCandidate> candidates) {
                IndexEntryCandidate first = candidates.First();
                IndexPageReference[] pageReferences = candidates
                    .Where(candidate => candidate.PageReference.HasValue)
                    .Select(candidate => candidate.PageReference!.Value)
                    .Distinct()
                    .OrderBy(reference => reference)
                    .ToArray();

                return new IndexEntry(first.Path, pageReferences, first.CrossReferenceText, first.EntryType);
            }

            internal WordIndexEntry ToPublicEntry(IndexSeparators separators) {
                return new WordIndexEntry(
                    Term,
                    Subterms,
                    PageNumbers,
                    PageReferences.Select(reference => reference.ToDisplayText(separators.PageRangeSeparator)).ToArray(),
                    CrossReferenceText,
                    EntryType,
                    separators.PageReferenceSeparator);
            }
        }

        private sealed class IndexNode {
            internal IndexNode(string name, int level) {
                Name = name;
                Level = level;
            }

            internal string Name { get; }

            internal int Level { get; }

            internal SortedDictionary<string, IndexNode> Children { get; } = new SortedDictionary<string, IndexNode>(StringComparer.OrdinalIgnoreCase);

            internal List<IndexEntry> Entries { get; } = new List<IndexEntry>();
        }

        private sealed class IndexPathComparer : IComparer<IReadOnlyList<string>> {
            internal static readonly IndexPathComparer Instance = new IndexPathComparer();

            public int Compare(IReadOnlyList<string>? x, IReadOnlyList<string>? y) {
                if (ReferenceEquals(x, y)) return 0;
                if (x == null) return -1;
                if (y == null) return 1;

                int count = Math.Min(x.Count, y.Count);
                for (int i = 0; i < count; i++) {
                    int comparison = string.Compare(x[i], y[i], StringComparison.OrdinalIgnoreCase);
                    if (comparison != 0) {
                        return comparison;
                    }
                }

                return x.Count.CompareTo(y.Count);
            }
        }
    }
}
