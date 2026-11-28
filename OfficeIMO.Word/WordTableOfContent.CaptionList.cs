using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
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
            if (ContentNode.IsNone) {
                throw new InvalidOperationException("Table of contents content block is missing.");
            }

            string normalizedSequence = NormalizeCaptionSequenceIdentifier(sequenceIdentifier);
            string instruction = GetCaptionListInstruction(normalizedSequence);
            ISet<int> pageNumberSuppressedLevels = GetTocPageNumberSuppressedLevels(instruction);
            string pageNumberSeparator = GetTocPageNumberSeparator(instruction);
            _ = ReadTocUpdateOnOpen();
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
            OfficeXmlNode body = WordFieldInventory.BodyNode(_document);
            if (body.IsNone) throw new InvalidOperationException("Document body is missing.");
            if (_document._ownedModel is { } model) WordOwnedEdits.PrepareBodyEdit(model);

            var entries = new List<CaptionListEntry>();
            var roots = WordFieldInventory.EnumerateStories(_document).ToArray();
            OfficeXmlNode control = SdtNode;
            Dictionary<OfficeXmlNode, int> paragraphPages = BuildParagraphPageMap(body, control);
            Dictionary<int, int> footnotePages = BuildNoteAnchorPageMap(body, paragraphPages, "footnoteReference");
            Dictionary<int, int> endnotePages = BuildNoteAnchorPageMap(body, paragraphPages, "endnoteReference");
            var bookmarks = new CaptionBookmarks(roots);

            skippedCaptionCount = 0;
            foreach (WordFieldInventory.NodeFieldRoot root in roots) {
                foreach (OfficeXmlNode paragraph in EnumerateTocSourceParagraphs(root.Root, control)) {
                    int page = GetParagraphPageNumber(paragraph, paragraphPages, footnotePages, endnotePages);
                    ProcessCaptionParagraph(paragraph, sequenceIdentifier, entries, bookmarks, page, ref skippedCaptionCount);
                }
            }
            return entries;
        }

        private void ProcessCaptionParagraph(OfficeXmlNode paragraph, string sequenceIdentifier, List<CaptionListEntry> entries,
            CaptionBookmarks bookmarks, int pageNumber, ref int skippedCaptionCount) {
            bool containsOwnCaption = ContainsSequenceField(paragraph, sequenceIdentifier, includeTextBoxContent: false);
            if (containsOwnCaption) {
                AddCaptionEntry(paragraph, sequenceIdentifier, entries, bookmarks, pageNumber, ref skippedCaptionCount,
                    includeTextBoxBookmarks: false);
            }

            HashSet<string>? seenCaptionText = null;
            foreach (OfficeXmlNode textBox in EnumerateTocElements(paragraph, "txbxContent")) {
                foreach (OfficeXmlNode caption in EnumerateTocElements(textBox, "p")) {
                    if (!ContainsSequenceField(caption, sequenceIdentifier)) continue;
                    seenCaptionText ??= new HashSet<string>(StringComparer.Ordinal);
                    AddCaptionEntry(caption, sequenceIdentifier, entries, bookmarks, pageNumber, ref skippedCaptionCount,
                        seenCaptionText);
                }
            }
            if (seenCaptionText == null && !containsOwnCaption && ContainsSequenceField(paragraph, sequenceIdentifier)) {
                AddCaptionEntry(paragraph, sequenceIdentifier, entries, bookmarks, pageNumber, ref skippedCaptionCount);
            }
        }

        private void AddCaptionEntry(OfficeXmlNode paragraph, string sequenceIdentifier, List<CaptionListEntry> entries,
            CaptionBookmarks bookmarks, int pageNumber, ref int skippedCaptionCount, HashSet<string>? seenCaptionText = null,
            bool includeTextBoxBookmarks = true) {
            string captionText = GetParagraphText(paragraph).Trim();
            if (captionText.Length == 0) {
                skippedCaptionCount++;
                return;
            }
            if (seenCaptionText != null && !seenCaptionText.Add(captionText)) return;

            string bookmarkName = EnsureCaptionBookmark(paragraph, bookmarks, sequenceIdentifier, entries.Count, includeTextBoxBookmarks);
            entries.Add(new CaptionListEntry(captionText, pageNumber, bookmarkName));
        }

        private void ReplaceCaptionListContent(string instruction, string sequenceIdentifier, string title, IReadOnlyList<CaptionListEntry> entries,
            ISet<int> pageNumberSuppressedLevels, string pageNumberSeparator) {
            const int W = TocNamespace;
            string noContentText = GetNoContentText("No " + sequenceIdentifier + " entries found.");
            OfficeXmlNode content = ResetTocContentWithTitle(title);
            var paragraphs = new List<OfficeXmlNode>(entries.Count);
            foreach (CaptionListEntry entry in entries) {
                OfficeXmlNode paragraph = content.CreateElement(W, "p");
                AppendTocChild(AppendTocChild(paragraph, "pPr"), "pStyle").SetAttribute(W, "val", "TOC1");
                OfficeXmlNode link = AppendTocChild(paragraph, "hyperlink");
                link.SetAttribute(W, "anchor", entry.BookmarkName);
                link.SetAttribute(W, "history", "true");
                link.Append(CreateTocTextRun(link, "t", entry.Text, preserveSpace: true));
                if (!pageNumberSuppressedLevels.Contains(1)) {
                    if (pageNumberSeparator == "\t") {
                        AppendTocChild(AppendTocChild(link, "r"), "tab");
                    } else if (pageNumberSeparator.Length > 0) {
                        link.Append(CreateTocTextRun(link, "t", pageNumberSeparator, preserveSpace: true));
                    }
                    link.Append(CreateTocTextRun(link, "t", entry.PageNumber.ToString(CultureInfo.InvariantCulture), preserveSpace: false));
                }
                paragraphs.Add(paragraph);
            }
            AppendGeneratedFieldResult(content, instruction, paragraphs, noContentText);
        }

        private static bool ContainsSequenceField(OfficeXmlNode paragraph, string sequenceIdentifier, bool includeTextBoxContent = true) {
            const int W = TocNamespace;
            int depth = 0;
            System.Text.StringBuilder? complex = null;
            foreach (OfficeXmlNode element in WordFieldInventory.EnumerateFieldContent(paragraph)) {
                if (element.NamespaceId != W ||
                    !includeTextBoxContent && element.AncestorOrSelf(W, "txbxContent").Exists) continue;
                if (element.Is(W, "fldSimple") && IsSequenceInstruction(element.Attribute(W, "instr"), sequenceIdentifier)) {
                    return true;
                }
                if (element.Is(W, "fldChar")) {
                    if (element.Attribute(W, "fldCharType") == "begin") {
                        if (depth++ == 0) complex?.Clear();
                    } else if (element.Attribute(W, "fldCharType") == "end" && depth > 0 && --depth == 0 &&
                               complex?.Length > 0 && IsSequenceInstruction(complex.ToString(), sequenceIdentifier)) {
                        return true;
                    }
                } else if (element.Is(W, "instrText")) {
                    string instruction = element.InnerText;
                    if (IsSequenceInstruction(instruction, sequenceIdentifier)) return true;
                    if (depth > 0) (complex ??= new System.Text.StringBuilder()).Append(instruction);
                }
            }
            return false;
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

        private string EnsureCaptionBookmark(OfficeXmlNode paragraph, CaptionBookmarks bookmarks, string sequenceIdentifier,
            int captionIndex, bool includeTextBoxBookmarks) {
            const int W = TocNamespace;
            foreach (OfficeXmlNode bookmark in EnumerateTocElements(paragraph, "bookmarkStart")) {
                if ((includeTextBoxBookmarks || bookmark.AncestorOrSelf(W, "txbxContent").IsNone) &&
                    bookmark.Attribute(W, "name") is { } name && !string.IsNullOrWhiteSpace(name)) return name;
            }

            string safeSequence = new string(sequenceIdentifier.Select(character => char.IsLetterOrDigit(character) ? character : '_').ToArray());
            string prefix = "_OfficeIMO_Caption_" + safeSequence + "_" + captionIndex.ToString(CultureInfo.InvariantCulture) + "_";
            string bookmarkName = bookmarks.AllocateName(prefix);
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

        private sealed class CaptionBookmarks {
            private readonly HashSet<string> _names = new(StringComparer.Ordinal);
            private int _maxId = -1;
            private int _nextNameSuffix;

            internal CaptionBookmarks(IReadOnlyList<WordFieldInventory.NodeFieldRoot> stories) {
                const int W = TocNamespace;
                foreach (WordFieldInventory.NodeFieldRoot story in stories) {
                    foreach (OfficeXmlNode bookmark in EnumerateTocElements(story.Root, "bookmarkStart")) {
                        if (bookmark.Attribute(W, "name") is { } name && !string.IsNullOrWhiteSpace(name)) _names.Add(name);
                    }
                }
                _nextNameSuffix = _names.Count;
                foreach (WordFieldInventory.NodeFieldRoot story in stories) {
                    foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(story.Root, reserveInactiveCompatibilityIds: true)) {
                        bool isStart = node.Is(W, "bookmarkStart");
                        if (isStart && node.Attribute(W, "name") is { } name && !string.IsNullOrWhiteSpace(name)) _names.Add(name);
                        if ((isStart || node.Is(W, "bookmarkEnd")) &&
                            OfficeXmlViewValues.ParseInt32(node.Attribute(W, "id")) is int id) _maxId = Math.Max(_maxId, id);
                    }
                }
            }

            internal string AllocateName(string prefix) {
                int suffix = _nextNameSuffix++;
                string name;
                do {
                    name = prefix + suffix.ToString(CultureInfo.InvariantCulture);
                    suffix++;
                } while (!_names.Add(name));
                return name;
            }

            internal int AllocateId() {
                if (_maxId == int.MaxValue) throw new InvalidOperationException("The document has exhausted the available bookmark identifier range.");
                return ++_maxId;
            }
        }

        private static string NormalizeCaptionSequenceIdentifier(string sequenceIdentifier) {
            ArgumentException.ThrowIfNullOrWhiteSpace(sequenceIdentifier);

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
