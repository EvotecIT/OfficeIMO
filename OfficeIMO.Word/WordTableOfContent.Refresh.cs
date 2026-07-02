using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
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
        /// This deterministic refresh uses explicit page breaks to estimate page numbers. It preserves a TOC field
        /// paragraph so Word can still perform a full layout-aware update when the document is opened.
        /// </remarks>
        /// <returns>A report describing the generated entries and skipped headings.</returns>
        public WordTableOfContentRefreshReport RefreshEntries() {
            if (_sdtBlock.SdtContentBlock == null) {
                throw new InvalidOperationException("Table of contents content block is missing.");
            }

            string instruction = GetTocInstruction();
            string noContentText = GetNoContentText();
            string? tcEntryTypeFilter = GetTocEntryTypeFilter(instruction);
            string? bookmarkScope = GetTocBookmarkScope(instruction);
            ISet<int> pageNumberSuppressedLevels = GetTocPageNumberSuppressedLevels(instruction);
            string pageNumberSeparator = GetTocPageNumberSeparator(instruction);
            IReadOnlyDictionary<string, int> customStyleLevels = GetTocCustomStyleLevels(instruction);
            TocSourceOptions sourceOptions = GetTocSourceOptions(instruction, customStyleLevels);
            var headings = CollectHeadingEntries(tcEntryTypeFilter, customStyleLevels, bookmarkScope, sourceOptions).ToArray();
            var included = headings
                .Where(heading => !heading.ApplyLevelRange || (heading.Level >= MinLevel && heading.Level <= MaxLevel))
                .ToArray();
            int skipped = headings.Length - included.Length;

            ReplaceContent(instruction, noContentText, included, pageNumberSuppressedLevels, pageNumberSeparator);
            QueueUpdateOnOpen(force: true);

            return new WordTableOfContentRefreshReport(
                included.Select(heading => heading.ToPublicEntry()).ToArray(),
                skipped,
                "Estimated from explicit page breaks; Word may recalculate layout on open.");
        }

        private IReadOnlyList<TocHeadingEntry> CollectHeadingEntries(
            string? tcEntryTypeFilter,
            IReadOnlyDictionary<string, int> customStyleLevels,
            string? bookmarkScope,
            TocSourceOptions sourceOptions) {
            Body body = _document._wordprocessingDocument.MainDocumentPart?.Document?.Body
                ?? throw new InvalidOperationException("Document body is missing.");

            var entries = new List<TocHeadingEntry>();
            BookmarkRange? bookmarkRange = bookmarkScope == null
                ? null
                : FindBookmarkRange(body, bookmarkScope);
            var existingBookmarkNames = new HashSet<string>(
                body.Descendants<BookmarkStart>()
                    .Select(bookmark => bookmark.Name?.Value)
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Cast<string>(),
                StringComparer.Ordinal);

            int pageNumber = 1;
            int headingIndex = 0;

            foreach (OpenXmlElement child in body.ChildElements) {
                if (ReferenceEquals(child, _sdtBlock)) {
                    continue;
                }

                if (child is Paragraph paragraph) {
                    ProcessHeadingParagraphWithTextBoxes(paragraph, tcEntryTypeFilter, customStyleLevels, bookmarkScope, bookmarkRange, entries, existingBookmarkNames, sourceOptions, ref pageNumber, ref headingIndex);
                    continue;
                }

                if (child is Table table) {
                    foreach (Paragraph nestedParagraph in table.Descendants<Paragraph>()) {
                        if (nestedParagraph.Ancestors<TextBoxContent>().Any()) {
                            continue;
                        }

                        ProcessHeadingParagraphWithTextBoxes(nestedParagraph, tcEntryTypeFilter, customStyleLevels, bookmarkScope, bookmarkRange, entries, existingBookmarkNames, sourceOptions, ref pageNumber, ref headingIndex);
                    }

                    continue;
                }

                if (child is SdtBlock sdtBlock) {
                    foreach (Paragraph nestedParagraph in sdtBlock.Descendants<Paragraph>()) {
                        if (nestedParagraph.Ancestors<TextBoxContent>().Any()) {
                            continue;
                        }

                        ProcessHeadingParagraphWithTextBoxes(nestedParagraph, tcEntryTypeFilter, customStyleLevels, bookmarkScope, bookmarkRange, entries, existingBookmarkNames, sourceOptions, ref pageNumber, ref headingIndex);
                    }
                }
            }

            return entries;
        }

        private void ProcessHeadingParagraphWithTextBoxes(
            Paragraph paragraph,
            string? tcEntryTypeFilter,
            IReadOnlyDictionary<string, int> customStyleLevels,
            string? bookmarkScope,
            BookmarkRange? bookmarkRange,
            List<TocHeadingEntry> entries,
            HashSet<string> existingBookmarkNames,
            TocSourceOptions sourceOptions,
            ref int pageNumber,
            ref int headingIndex) {
            if (paragraph.ParagraphProperties?.PageBreakBefore != null) {
                pageNumber++;
            }

            foreach (Paragraph textBoxHeadingParagraph in GetBodyTextBoxHeadingParagraphs(paragraph, customStyleLevels, sourceOptions)) {
                ProcessHeadingParagraph(textBoxHeadingParagraph, tcEntryTypeFilter, customStyleLevels, bookmarkScope, bookmarkRange, entries, existingBookmarkNames, sourceOptions, ref pageNumber, ref headingIndex, countPageBreakBefore: false, countTrailingBreaks: false);
            }

            ProcessHeadingParagraph(paragraph, tcEntryTypeFilter, customStyleLevels, bookmarkScope, bookmarkRange, entries, existingBookmarkNames, sourceOptions, ref pageNumber, ref headingIndex, countPageBreakBefore: false, countTrailingBreaks: false);
            CountParagraphPageBreaks(paragraph, ref pageNumber);
        }

        private static IReadOnlyList<Paragraph> GetBodyTextBoxHeadingParagraphs(
            Paragraph paragraph,
            IReadOnlyDictionary<string, int> customStyleLevels,
            TocSourceOptions sourceOptions) {
            if (paragraph.Ancestors<TextBoxContent>().Any()) {
                return Array.Empty<Paragraph>();
            }

            var headings = new List<Paragraph>();
            var seenWordParagraphIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (Paragraph candidate in paragraph.Descendants<TextBoxContent>()
                .SelectMany(textBoxContent => textBoxContent.Descendants<Paragraph>())) {
                if (GetHeadingLevel(candidate, customStyleLevels, sourceOptions) == null) {
                    continue;
                }

                string? paragraphId = candidate.ParagraphId?.Value;
                if (!string.IsNullOrWhiteSpace(paragraphId) && !seenWordParagraphIds.Add(paragraphId!)) {
                    continue;
                }

                headings.Add(candidate);
            }

            return headings;
        }

        private void ProcessHeadingParagraph(
            Paragraph paragraph,
            string? tcEntryTypeFilter,
            IReadOnlyDictionary<string, int> customStyleLevels,
            string? bookmarkScope,
            BookmarkRange? bookmarkRange,
            List<TocHeadingEntry> entries,
            HashSet<string> existingBookmarkNames,
            TocSourceOptions sourceOptions,
            ref int pageNumber,
            ref int headingIndex,
            bool countPageBreakBefore = true,
            bool countTrailingBreaks = true) {
            if (countPageBreakBefore && paragraph.ParagraphProperties?.PageBreakBefore != null) {
                pageNumber++;
            }

            bool isInsideScope = bookmarkScope == null || (bookmarkRange.HasValue && IsParagraphInsideBookmarkRange(paragraph, bookmarkRange.Value));
            if (isInsideScope) {
                (int Level, bool ApplyLevelRange)? headingLevel = GetHeadingLevel(paragraph, customStyleLevels, sourceOptions);
                if (headingLevel != null) {
                    string headingText = GetParagraphText(paragraph);
                    if (!string.IsNullOrWhiteSpace(headingText)) {
                        RemoveImportedTableTocBookmarks(paragraph, existingBookmarkNames);
                        string bookmarkName = EnsureBookmark(paragraph, existingBookmarkNames, headingIndex);
                        entries.Add(new TocHeadingEntry(headingText, headingLevel.Value.Level, pageNumber, bookmarkName, headingLevel.Value.ApplyLevelRange));
                        headingIndex++;
                    }
                }

                if (sourceOptions.IncludeTocEntryFields) {
                    foreach (string instruction in EnumerateTocEntryInstructions(paragraph)) {
                        if (TryCreateTocEntry(instruction, tcEntryTypeFilter, paragraph, existingBookmarkNames, headingIndex, pageNumber, out TocHeadingEntry? entry)) {
                            entries.Add(entry!);
                            headingIndex++;
                        }
                    }
                }
            }

            if (!countTrailingBreaks) {
                return;
            }

            CountParagraphPageBreaks(paragraph, ref pageNumber);
        }

        private static void CountParagraphPageBreaks(Paragraph paragraph, ref int pageNumber) {
            pageNumber += paragraph.Descendants<Break>().Count(documentBreak => documentBreak.Type?.Value == BreakValues.Page);

            if (StartsNewPage(paragraph.ParagraphProperties?.SectionProperties)) {
                pageNumber++;
            }
        }

        private static IEnumerable<string> EnumerateTocEntryInstructions(Paragraph paragraph) {
            foreach (SimpleField simpleField in paragraph.Descendants<SimpleField>()) {
                string instruction = simpleField.Instruction?.Value ?? string.Empty;
                if (IsTocEntryInstruction(instruction)) {
                    yield return instruction;
                }
            }

            List<string>? currentComplexField = null;
            foreach (OpenXmlElement element in paragraph.Descendants()) {
                if (element is FieldChar fieldChar) {
                    if (fieldChar.FieldCharType?.Value == FieldCharValues.Begin) {
                        currentComplexField = new List<string>();
                    } else if (fieldChar.FieldCharType?.Value == FieldCharValues.End && currentComplexField != null) {
                        string instruction = string.Concat(currentComplexField);
                        if (IsTocEntryInstruction(instruction)) {
                            yield return instruction;
                        }

                        currentComplexField = null;
                    }

                    continue;
                }

                if (element is FieldCode fieldCode) {
                    if (currentComplexField != null) {
                        currentComplexField.Add(fieldCode.Text ?? string.Empty);
                    } else {
                        string instruction = fieldCode.Text ?? string.Empty;
                        if (IsTocEntryInstruction(instruction)) {
                            yield return instruction;
                        }
                    }
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
            Paragraph paragraph,
            HashSet<string> existingBookmarkNames,
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

            string bookmarkName = EnsureBookmark(paragraph, existingBookmarkNames, entryIndex);
            entry = new TocHeadingEntry(text, level, pageNumber, bookmarkName, applyLevelRange: false);
            return true;
        }

        private static bool StartsNewPage(SectionProperties? sectionProperties) {
            if (sectionProperties == null) {
                return false;
            }

            SectionType? sectionType = sectionProperties.GetFirstChild<SectionType>();
            SectionMarkValues? value = sectionType?.Val?.Value;
            return value == null ||
                   value == SectionMarkValues.NextPage ||
                   value == SectionMarkValues.OddPage ||
                   value == SectionMarkValues.EvenPage;
        }

        private void ReplaceContent(string instruction, string noContentText, IReadOnlyList<TocHeadingEntry> entries, ISet<int> pageNumberSuppressedLevels, string pageNumberSeparator) {
            SdtContentBlock content = _sdtBlock.SdtContentBlock
                ?? throw new InvalidOperationException("Table of contents content block is missing.");

            Paragraph titleParagraph = GetTitleParagraph() ?? CreateTitleParagraph("Table of Contents");
            RemoveIfAttached(titleParagraph);
            content.RemoveAllChildren();
            content.Append(titleParagraph);

            content.Append(CreateTocFieldParagraph(instruction, entries.Count == 0 ? noContentText : string.Empty));

            foreach (TocHeadingEntry entry in entries) {
                content.Append(CreateEntryParagraph(entry, !pageNumberSuppressedLevels.Contains(entry.Level), pageNumberSeparator));
            }
        }

        private Paragraph? GetTitleParagraph() {
            SdtContentBlock? content = _sdtBlock.SdtContentBlock;
            if (content == null) {
                return null;
            }

            return content.ChildElements
                .OfType<Paragraph>()
                .FirstOrDefault(paragraph =>
                    !paragraph.Descendants<SimpleField>().Any() &&
                    !paragraph.Descendants<FieldCode>().Any() &&
                    !IsGeneratedResultParagraph(paragraph));
        }

        private static bool IsGeneratedResultParagraph(Paragraph paragraph) {
            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) {
                return false;
            }

            return Regex.IsMatch(styleId!, "^(TOC|Index)[1-9]$", RegexOptions.IgnoreCase) ||
                   string.Equals(styleId, "TableofFigures", StringComparison.OrdinalIgnoreCase);
        }

        private string GetTocInstruction() {
            foreach (SimpleField simpleField in _sdtBlock.Descendants<SimpleField>()) {
                string? instruction = simpleField.Instruction?.Value ?? simpleField.Instruction;
                string instructionText = instruction ?? string.Empty;
                if (IsTocInstruction(instructionText)) {
                    return instructionText;
                }
            }

            foreach (Paragraph paragraph in _sdtBlock.Descendants<Paragraph>()) {
                foreach (string instruction in EnumerateComplexFieldInstructions(paragraph)) {
                    if (IsTocInstruction(instruction)) {
                        return instruction;
                    }
                }
            }

            foreach (FieldCode fieldCode in _sdtBlock.Descendants<FieldCode>()) {
                string instruction = fieldCode.Text ?? string.Empty;
                if (IsTocInstruction(instruction)) {
                    return instruction;
                }
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
                return new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            }

            var levels = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\t", StringComparison.OrdinalIgnoreCase))) {
                string value = TrimFieldArgument(fieldSwitch.Trim().Substring(2).Trim());
                AddCustomStyleLevelPairs(value, levels);
            }

            if (levels.Count == 0) {
                return levels;
            }

            AddCustomStyleAliases(levels);
            return levels;
        }

        private static void AddCustomStyleLevelPairs(string value, Dictionary<string, int> levels) {
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

                levels[styleName] = ClampLevel(parsedLevel);
            }
        }

        private void AddCustomStyleAliases(Dictionary<string, int> levels) {
            Styles? styles = _document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) {
                return;
            }

            foreach (Style style in styles.Elements<Style>().Where(style => style.Type?.Value == StyleValues.Paragraph)) {
                string? styleId = style.StyleId?.Value;
                string? styleName = style.StyleName?.Val?.Value;
                if (string.IsNullOrWhiteSpace(styleId) || string.IsNullOrWhiteSpace(styleName)) {
                    continue;
                }

                if (levels.TryGetValue(styleName!, out int levelFromName) && !levels.ContainsKey(styleId!)) {
                    levels[styleId!] = levelFromName;
                }

                if (levels.TryGetValue(styleId!, out int levelFromId) && !levels.ContainsKey(styleName!)) {
                    levels[styleName!] = levelFromId;
                }
            }
        }

        private string GetNoContentText() {
            string text = TextNoContent;
            return string.IsNullOrWhiteSpace(text) ? DefaultNoContentText : text;
        }

        private static Paragraph CreateTitleParagraph(string text) {
            return new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        private static Paragraph CreateTocFieldParagraph(string instruction, string resultText) {
            return new Paragraph(
                new SimpleField(
                    new Run(new Text(resultText) { Space = SpaceProcessingModeValues.Preserve })) {
                    Instruction = instruction,
                    Dirty = true
                });
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

        private static (int Level, bool ApplyLevelRange)? GetHeadingLevel(Paragraph paragraph, IReadOnlyDictionary<string, int> customStyleLevels, TocSourceOptions sourceOptions) {
            int? outlineLevel = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
            if (sourceOptions.IncludeOutlineLevels && outlineLevel >= 0 && outlineLevel <= 8) {
                return (outlineLevel.Value + 1, true);
            }

            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(styleId) &&
                sourceOptions.IncludeCustomStyles &&
                customStyleLevels.TryGetValue(styleId!, out int customLevel)) {
                return (customLevel, false);
            }

            if (string.IsNullOrWhiteSpace(styleId) || !sourceOptions.IncludeHeadingStyles) {
                return null;
            }

            Match match = Regex.Match(styleId, "^Heading(?<level>[1-9])$", RegexOptions.IgnoreCase);
            if (!match.Success || !int.TryParse(match.Groups["level"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int level)) {
                return null;
            }

            return (level, true);
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

        private static string GetParagraphText(Paragraph paragraph) {
            if (paragraph.Ancestors<TextBoxContent>().Any()) {
                return string.Concat(paragraph.Descendants<Text>().Select(text => text.Text));
            }

            return string.Concat(paragraph.Descendants<Text>()
                .Where(text => !text.Ancestors<TextBoxContent>().Any())
                .Select(text => text.Text));
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

        private string EnsureBookmark(Paragraph paragraph, HashSet<string> existingBookmarkNames, int headingIndex) {
            bool paragraphInTextBox = paragraph.Ancestors<TextBoxContent>().Any();
            BookmarkStart? existing = paragraph.Descendants<BookmarkStart>()
                .Where(bookmark => paragraphInTextBox || !bookmark.Ancestors<TextBoxContent>().Any())
                .FirstOrDefault(bookmark => !string.IsNullOrWhiteSpace(bookmark.Name?.Value));

            string? existingName = existing?.Name?.Value;
            if (existing != null && existingName != null && existingName.Trim().Length > 0) {
                return existingName;
            }

            string bookmarkName;
            do {
                bookmarkName = "_OfficeIMO_Toc_" + headingIndex.ToString(CultureInfo.InvariantCulture) + "_" + existingBookmarkNames.Count.ToString(CultureInfo.InvariantCulture);
            } while (!existingBookmarkNames.Add(bookmarkName));

            string bookmarkId = GetNextBookmarkId(paragraph).ToString(CultureInfo.InvariantCulture);
            var bookmarkStart = new BookmarkStart {
                Name = bookmarkName,
                Id = bookmarkId
            };
            var bookmarkEnd = new BookmarkEnd {
                Id = bookmarkId
            };

            ParagraphProperties? properties = paragraph.GetFirstChild<ParagraphProperties>();
            if (properties != null) {
                paragraph.InsertAfter(bookmarkStart, properties);
            } else {
                paragraph.PrependChild(bookmarkStart);
            }

            paragraph.Append(bookmarkEnd);
            return bookmarkName;
        }

        private static int GetNextBookmarkId(Paragraph paragraph) {
            Body? body = paragraph.Ancestors<Body>().FirstOrDefault();
            OpenXmlElement root = body != null ? body : paragraph;
            int maxId = 0;
            foreach (OpenXmlElement bookmark in root.Descendants<BookmarkStart>().Cast<OpenXmlElement>().Concat(root.Descendants<BookmarkEnd>())) {
                string? value = bookmark switch {
                    BookmarkStart start => start.Id?.Value,
                    BookmarkEnd end => end.Id?.Value,
                    _ => null
                };
                if (int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int id) && id > maxId) {
                    maxId = id;
                }
            }

            return maxId + 1;
        }

        private static void RemoveImportedTableTocBookmarks(Paragraph paragraph, HashSet<string> existingBookmarkNames) {
            Table? table = paragraph.Ancestors<Table>().FirstOrDefault();
            if (table == null) {
                return;
            }

            foreach (BookmarkStart bookmarkStart in paragraph.Descendants<BookmarkStart>().ToArray()) {
                string? bookmarkName = bookmarkStart.Name?.Value;
                string? bookmarkId = bookmarkStart.Id?.Value;
                if (string.IsNullOrWhiteSpace(bookmarkName) ||
                    string.IsNullOrWhiteSpace(bookmarkId) ||
                    !bookmarkName!.StartsWith("_Toc", StringComparison.Ordinal)) {
                    continue;
                }

                BookmarkEnd? bookmarkEnd = table.Descendants<BookmarkEnd>()
                    .FirstOrDefault(end => end.Id?.Value == bookmarkId);
                if (bookmarkEnd == null) {
                    continue;
                }

                bookmarkStart.Remove();
                bookmarkEnd.Remove();
                existingBookmarkNames.Remove(bookmarkName);
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
