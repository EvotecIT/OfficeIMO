using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word {
    public partial class WordTableOfContent {
        private const string DefaultCaptionListPageNumberMode = "Estimated from explicit page breaks; Word may recalculate layout on open.";

        /// <summary>
        /// Regenerates visible list entries from generated caption paragraphs that contain a matching SEQ field.
        /// </summary>
        /// <param name="sequenceIdentifier">Caption sequence identifier, for example <c>Figure</c> or <c>Table</c>.</param>
        /// <param name="title">Optional title shown above the generated list.</param>
        /// <returns>A report describing the generated caption-list entries.</returns>
        public WordCaptionListRefreshReport RefreshCaptionList(string sequenceIdentifier, string? title = null) {
            if (_sdtBlock.SdtContentBlock == null) {
                throw new InvalidOperationException("Table of contents content block is missing.");
            }

            string normalizedSequence = NormalizeCaptionSequenceIdentifier(sequenceIdentifier);
            string instruction = GetCaptionListInstruction(normalizedSequence);
            ISet<int> pageNumberSuppressedLevels = GetTocPageNumberSuppressedLevels(instruction);
            string pageNumberSeparator = GetTocPageNumberSeparator(instruction);
            var entries = CollectCaptionEntries(normalizedSequence, out int skippedCaptionCount).ToArray();
            string listTitle = string.IsNullOrWhiteSpace(title) ? "List of " + normalizedSequence + "s" : title!.Trim();

            ReplaceCaptionListContent(instruction, normalizedSequence, listTitle, entries, pageNumberSuppressedLevels, pageNumberSeparator);
            QueueUpdateOnOpen(force: true);

            return new WordCaptionListRefreshReport(
                normalizedSequence,
                entries.Select(entry => entry.ToPublicEntry(normalizedSequence)).ToArray(),
                skippedCaptionCount,
                DefaultCaptionListPageNumberMode);
        }

        /// <summary>
        /// Regenerates a generated list of figures from caption paragraphs that contain <c>SEQ Figure</c> fields.
        /// </summary>
        /// <returns>A report describing the generated figure-list entries.</returns>
        public WordCaptionListRefreshReport RefreshListOfFigures() {
            return RefreshCaptionList("Figure", "List of Figures");
        }

        /// <summary>
        /// Regenerates a generated list of tables from caption paragraphs that contain <c>SEQ Table</c> fields.
        /// </summary>
        /// <returns>A report describing the generated table-list entries.</returns>
        public WordCaptionListRefreshReport RefreshListOfTables() {
            return RefreshCaptionList("Table", "List of Tables");
        }

        private IReadOnlyList<CaptionListEntry> CollectCaptionEntries(string sequenceIdentifier, out int skippedCaptionCount) {
            MainDocumentPart mainPart = _document._wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("Main document part is missing.");

            var entries = new List<CaptionListEntry>();
            var roots = WordFieldInventory.EnumerateFieldRoots(mainPart).ToArray();
            var existingBookmarkNames = new HashSet<string>(
                roots
                    .SelectMany(root => root.Root.Descendants<BookmarkStart>())
                    .Select(bookmark => bookmark.Name?.Value)
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Cast<string>(),
                StringComparer.Ordinal);

            int pageNumber = 1;
            int captionIndex = 0;
            skippedCaptionCount = 0;

            foreach (WordFieldInventory.FieldRoot root in roots) {
                foreach (OpenXmlElement child in root.Root.ChildElements) {
                    if (ReferenceEquals(child, _sdtBlock)) {
                        continue;
                    }

                    if (child is Paragraph paragraph) {
                        ProcessCaptionParagraph(paragraph, sequenceIdentifier, entries, existingBookmarkNames, ref pageNumber, ref captionIndex, ref skippedCaptionCount);
                        continue;
                    }

                    if (child is Table || child is SdtBlock) {
                        ProcessNestedCaptionParagraphs(child, sequenceIdentifier, entries, existingBookmarkNames, ref pageNumber, ref captionIndex, ref skippedCaptionCount);
                        continue;
                    }

                    if (child is Footnote footnote && IsVisibleNote(footnote)) {
                        ProcessNestedCaptionParagraphs(footnote, sequenceIdentifier, entries, existingBookmarkNames, ref pageNumber, ref captionIndex, ref skippedCaptionCount);
                        continue;
                    }

                    if (child is Endnote endnote && IsVisibleNote(endnote)) {
                        ProcessNestedCaptionParagraphs(endnote, sequenceIdentifier, entries, existingBookmarkNames, ref pageNumber, ref captionIndex, ref skippedCaptionCount);
                    }
                }
            }

            return entries;
        }

        private void ProcessNestedCaptionParagraphs(
            OpenXmlElement container,
            string sequenceIdentifier,
            List<CaptionListEntry> entries,
            HashSet<string> existingBookmarkNames,
            ref int pageNumber,
            ref int captionIndex,
            ref int skippedCaptionCount) {
            foreach (Paragraph nestedParagraph in container.Descendants<Paragraph>()) {
                if (nestedParagraph.Ancestors<TextBoxContent>().Any()) {
                    continue;
                }

                ProcessCaptionParagraph(nestedParagraph, sequenceIdentifier, entries, existingBookmarkNames, ref pageNumber, ref captionIndex, ref skippedCaptionCount);
            }
        }

        private void ProcessCaptionParagraph(
            Paragraph paragraph,
            string sequenceIdentifier,
            List<CaptionListEntry> entries,
            HashSet<string> existingBookmarkNames,
            ref int pageNumber,
            ref int captionIndex,
            ref int skippedCaptionCount) {
            if (paragraph.ParagraphProperties?.PageBreakBefore != null) {
                pageNumber++;
            }

            IReadOnlyList<Paragraph> textBoxCaptionParagraphs = GetTextBoxCaptionParagraphs(paragraph, sequenceIdentifier);
            if (textBoxCaptionParagraphs.Count > 0) {
                var seenCaptionText = new HashSet<string>(StringComparer.Ordinal);
                foreach (Paragraph textBoxCaptionParagraph in textBoxCaptionParagraphs) {
                    AddCaptionEntry(
                        textBoxCaptionParagraph,
                        sequenceIdentifier,
                        entries,
                        existingBookmarkNames,
                        pageNumber,
                        ref captionIndex,
                        ref skippedCaptionCount,
                        seenCaptionText);
                }
            } else if (ContainsSequenceField(paragraph, sequenceIdentifier)) {
                AddCaptionEntry(
                    paragraph,
                    sequenceIdentifier,
                    entries,
                    existingBookmarkNames,
                    pageNumber,
                    ref captionIndex,
                    ref skippedCaptionCount);
            }

            pageNumber += paragraph.Descendants<Break>().Count(documentBreak => documentBreak.Type?.Value == BreakValues.Page);

            if (StartsNewPage(paragraph.ParagraphProperties?.SectionProperties)) {
                pageNumber++;
            }
        }

        private void AddCaptionEntry(
            Paragraph paragraph,
            string sequenceIdentifier,
            List<CaptionListEntry> entries,
            HashSet<string> existingBookmarkNames,
            int pageNumber,
            ref int captionIndex,
            ref int skippedCaptionCount,
            HashSet<string>? seenCaptionText = null) {
            string captionText = GetParagraphText(paragraph).Trim();
            if (captionText.Length == 0) {
                skippedCaptionCount++;
                return;
            }

            if (seenCaptionText != null && !seenCaptionText.Add(captionText)) {
                return;
            }

            string bookmarkName = EnsureCaptionBookmark(paragraph, existingBookmarkNames, sequenceIdentifier, captionIndex);
            entries.Add(new CaptionListEntry(captionText, pageNumber, bookmarkName));
            captionIndex++;
        }

        private static IReadOnlyList<Paragraph> GetTextBoxCaptionParagraphs(Paragraph paragraph, string sequenceIdentifier) {
            var captionParagraphs = new List<Paragraph>();
            foreach (TextBoxContent textBoxContent in paragraph.Descendants<TextBoxContent>()) {
                captionParagraphs.AddRange(textBoxContent.Descendants<Paragraph>()
                    .Where(candidate => ContainsSequenceField(candidate, sequenceIdentifier)));
            }

            return captionParagraphs;
        }

        private void ReplaceCaptionListContent(string instruction, string sequenceIdentifier, string title, IReadOnlyList<CaptionListEntry> entries, ISet<int> pageNumberSuppressedLevels, string pageNumberSeparator) {
            SdtContentBlock content = _sdtBlock.SdtContentBlock
                ?? throw new InvalidOperationException("Table of contents content block is missing.");

            Paragraph titleParagraph = GetTitleParagraph() ?? CreateTitleParagraph(title);
            SetParagraphText(titleParagraph, title);
            RemoveIfAttached(titleParagraph);

            content.RemoveAllChildren();
            content.Append(titleParagraph);
            content.Append(CreateTocFieldParagraph(instruction, entries.Count == 0 ? "No " + sequenceIdentifier + " entries found." : string.Empty));

            foreach (CaptionListEntry entry in entries) {
                content.Append(CreateEntryParagraph(new TocHeadingEntry(entry.Text, 1, entry.PageNumber, entry.BookmarkName), !pageNumberSuppressedLevels.Contains(1), pageNumberSeparator));
            }
        }

        private static bool ContainsSequenceField(Paragraph paragraph, string sequenceIdentifier) {
            foreach (SimpleField simpleField in paragraph.Descendants<SimpleField>()) {
                if (IsSequenceInstruction(simpleField.Instruction?.Value ?? simpleField.Instruction, sequenceIdentifier)) {
                    return true;
                }
            }

            foreach (FieldCode fieldCode in paragraph.Descendants<FieldCode>()) {
                if (IsSequenceInstruction(fieldCode.Text, sequenceIdentifier)) {
                    return true;
                }
            }

            string combinedInstruction = string.Concat(paragraph.Descendants<FieldCode>().Select(fieldCode => fieldCode.Text));
            return IsSequenceInstruction(combinedInstruction, sequenceIdentifier);
        }

        private static bool IsSequenceInstruction(string? instruction, string sequenceIdentifier) {
            if (string.IsNullOrWhiteSpace(instruction)) {
                return false;
            }

            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction!);
            return parsed.FieldType == WordFieldType.Seq &&
                parsed.Instructions.Count > 0 &&
                string.Equals(TrimCaptionIdentifier(parsed.Instructions[0]), sequenceIdentifier, StringComparison.OrdinalIgnoreCase);
        }

        private string EnsureCaptionBookmark(Paragraph paragraph, HashSet<string> existingBookmarkNames, string sequenceIdentifier, int captionIndex) {
            BookmarkStart? existing = paragraph.Descendants<BookmarkStart>()
                .FirstOrDefault(bookmark => !string.IsNullOrWhiteSpace(bookmark.Name?.Value));

            string? existingName = existing?.Name?.Value;
            if (existingName != null && existingName.Trim().Length > 0) {
                return existingName;
            }

            string safeSequence = new string(sequenceIdentifier.Select(character => char.IsLetterOrDigit(character) ? character : '_').ToArray());
            string bookmarkName;
            do {
                bookmarkName = "_OfficeIMO_Caption_" + safeSequence + "_" + captionIndex.ToString(CultureInfo.InvariantCulture) + "_" + existingBookmarkNames.Count.ToString(CultureInfo.InvariantCulture);
            } while (!existingBookmarkNames.Add(bookmarkName));

            string bookmarkId = GetNextCaptionBookmarkId().ToString(CultureInfo.InvariantCulture);
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

        private int GetNextCaptionBookmarkId() {
            MainDocumentPart? mainPart = _document._wordprocessingDocument.MainDocumentPart;
            if (mainPart == null) {
                return _document.BookmarkId;
            }

            int maxId = -1;
            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (OpenXmlElement bookmark in root.Root.Descendants<BookmarkStart>().Cast<OpenXmlElement>().Concat(root.Root.Descendants<BookmarkEnd>())) {
                    string? value = bookmark switch {
                        BookmarkStart start => start.Id?.Value,
                        BookmarkEnd end => end.Id?.Value,
                        _ => null
                    };

                    if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)) {
                        maxId = Math.Max(maxId, id);
                    }
                }
            }

            return maxId + 1;
        }

        private static bool IsVisibleNote(Footnote footnote) {
            return footnote.Type == null || footnote.Type.Value == FootnoteEndnoteValues.Normal;
        }

        private static bool IsVisibleNote(Endnote endnote) {
            return endnote.Type == null || endnote.Type.Value == FootnoteEndnoteValues.Normal;
        }

        private static string NormalizeCaptionSequenceIdentifier(string sequenceIdentifier) {
            if (string.IsNullOrWhiteSpace(sequenceIdentifier)) {
                throw new ArgumentException("Caption sequence identifier cannot be empty.", nameof(sequenceIdentifier));
            }

            string normalized = TrimCaptionIdentifier(sequenceIdentifier);
            if (normalized.Length == 0 ||
                normalized.Any(char.IsWhiteSpace) ||
                normalized.IndexOf('\\') >= 0 ||
                normalized.IndexOf('"') >= 0) {
                throw new ArgumentException("Caption sequence identifier must be a single Word SEQ identifier without whitespace, quotes, or field switches.", nameof(sequenceIdentifier));
            }

            return normalized;
        }

        private static string TrimCaptionIdentifier(string value) {
            string trimmed = value.Trim();
            return trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"'
                ? trimmed.Substring(1, trimmed.Length - 2)
                : trimmed;
        }

        private static string CreateCaptionListInstruction(string sequenceIdentifier) {
            return " TOC \\h \\z \\c \"" + sequenceIdentifier + "\" ";
        }

        private string GetCaptionListInstruction(string sequenceIdentifier) {
            string instruction = GetTocInstruction();
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(instruction);
            if (parsed.FieldType == WordFieldType.TOC && parsed.Diagnostics.Count == 0) {
                foreach (string fieldSwitch in parsed.Switches.Where(item => item.Trim().StartsWith("\\c", StringComparison.OrdinalIgnoreCase))) {
                    string value = TrimFieldArgument(fieldSwitch.Trim().Substring(2).Trim());
                    if (string.Equals(value, sequenceIdentifier, StringComparison.OrdinalIgnoreCase)) {
                        return instruction;
                    }
                }
            }

            return CreateCaptionListInstruction(sequenceIdentifier);
        }

        private static void SetParagraphText(Paragraph paragraph, string value) {
            Text? text = paragraph.Descendants<Text>().FirstOrDefault();
            if (text == null) {
                paragraph.Append(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
                return;
            }

            text.Text = value;
            foreach (Text extraText in paragraph.Descendants<Text>().Skip(1)) {
                extraText.Text = string.Empty;
            }
        }

        private sealed class CaptionListEntry {
            internal CaptionListEntry(string text, int pageNumber, string bookmarkName) {
                Text = text;
                PageNumber = pageNumber;
                BookmarkName = bookmarkName;
            }

            internal string Text { get; }

            internal int PageNumber { get; }

            internal string BookmarkName { get; }

            internal WordCaptionListEntry ToPublicEntry(string sequenceIdentifier) {
                return new WordCaptionListEntry(sequenceIdentifier, Text, PageNumber, BookmarkName);
            }
        }
    }
}
