using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Marks matching paragraphs with hidden <c>XE</c> fields from in-memory concordance entries.
        /// </summary>
        /// <param name="entries">Concordance entries mapping search text to index text.</param>
        /// <param name="matchCase">Whether matching should be case-sensitive.</param>
        /// <param name="matchWholeWord">Whether matches require non-letter/digit boundaries around the search text.</param>
        /// <returns>A report describing the marks inserted and entries skipped.</returns>
        public WordIndexConcordanceReport MarkIndexEntries(IEnumerable<WordIndexConcordanceEntry> entries, bool matchCase = false, bool matchWholeWord = true) {
            if (entries == null) {
                throw new ArgumentNullException(nameof(entries));
            }

            Body body = _wordprocessingDocument.MainDocumentPart?.Document?.Body
                ?? throw new InvalidOperationException("Document body is missing.");

            WordIndexConcordanceEntry[] concordanceEntries = entries.ToArray();
            int skippedEntryCount = 0;
            int markedEntryCount = 0;
            var matchedParagraphs = new HashSet<Paragraph>();

            foreach (WordIndexConcordanceEntry entry in concordanceEntries) {
                if (entry == null || !CanUseIndexConcordanceText(entry.IndexText)) {
                    skippedEntryCount++;
                    continue;
                }

                string instruction = CreateIndexConcordanceInstruction(entry.IndexText);
                foreach (Paragraph paragraph in EnumerateConcordanceTargetParagraphs(body)) {
                    string paragraphText = GetParagraphText(paragraph);
                    if (!ContainsConcordanceMatch(paragraphText, entry.SearchText, matchCase, matchWholeWord)) {
                        continue;
                    }

                    if (HasIndexConcordanceInstruction(paragraph, instruction)) {
                        skippedEntryCount++;
                        continue;
                    }

                    paragraph.Append(CreateHiddenIndexConcordanceField(instruction));
                    markedEntryCount++;
                    matchedParagraphs.Add(paragraph);
                }
            }

            return new WordIndexConcordanceReport(
                concordanceEntries.Where(entry => entry != null && CanUseIndexConcordanceText(entry.IndexText)).ToArray(),
                markedEntryCount,
                matchedParagraphs.Count,
                skippedEntryCount,
                matchCase,
                matchWholeWord);
        }

        /// <summary>
        /// Reads a Word-style two-column concordance document and marks matching paragraphs with hidden <c>XE</c> fields.
        /// </summary>
        /// <param name="concordanceDocument">Concordance document whose first table column contains search text and second column contains index text.</param>
        /// <param name="matchCase">Whether matching should be case-sensitive.</param>
        /// <param name="matchWholeWord">Whether matches require non-letter/digit boundaries around the search text.</param>
        /// <returns>A report describing the marks inserted and entries skipped.</returns>
        public WordIndexConcordanceReport MarkIndexEntriesFromConcordance(WordDocument concordanceDocument, bool matchCase = false, bool matchWholeWord = true) {
            if (concordanceDocument == null) {
                throw new ArgumentNullException(nameof(concordanceDocument));
            }

            IReadOnlyList<WordIndexConcordanceEntry> entries = ReadIndexConcordanceEntries(concordanceDocument, out int skippedEntryCount);
            WordIndexConcordanceReport report = MarkIndexEntries(entries, matchCase, matchWholeWord);

            return new WordIndexConcordanceReport(
                report.Entries,
                report.MarkedEntryCount,
                report.MatchedParagraphCount,
                report.SkippedEntryCount + skippedEntryCount,
                report.MatchCase,
                report.MatchWholeWord);
        }

        /// <summary>
        /// Loads a Word-style two-column concordance document and marks matching paragraphs with hidden <c>XE</c> fields.
        /// </summary>
        /// <param name="concordancePath">Path to a concordance <c>.docx</c> file.</param>
        /// <param name="matchCase">Whether matching should be case-sensitive.</param>
        /// <param name="matchWholeWord">Whether matches require non-letter/digit boundaries around the search text.</param>
        /// <returns>A report describing the marks inserted and entries skipped.</returns>
        public WordIndexConcordanceReport MarkIndexEntriesFromConcordance(string concordancePath, bool matchCase = false, bool matchWholeWord = true) {
            if (string.IsNullOrWhiteSpace(concordancePath)) {
                throw new ArgumentException("Concordance path cannot be empty.", nameof(concordancePath));
            }

            using (WordDocument concordanceDocument = Load(concordancePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            })) {
                return MarkIndexEntriesFromConcordance(concordanceDocument, matchCase, matchWholeWord);
            }
        }

        private static IReadOnlyList<WordIndexConcordanceEntry> ReadIndexConcordanceEntries(WordDocument concordanceDocument, out int skippedEntryCount) {
            skippedEntryCount = 0;
            var entries = new List<WordIndexConcordanceEntry>();

            Body? body = concordanceDocument._wordprocessingDocument.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return entries;
            }

            foreach (TableRow row in body.Descendants<TableRow>()) {
                TableCell[] cells = row.Elements<TableCell>().Take(2).ToArray();
                if (cells.Length == 0) {
                    skippedEntryCount++;
                    continue;
                }

                string searchText = GetElementText(cells[0]).Trim();
                string indexText = cells.Length > 1 ? GetElementText(cells[1]).Trim() : searchText;
                if (indexText.Length == 0) {
                    indexText = searchText;
                }

                if (searchText.Length == 0 || indexText.Length == 0) {
                    skippedEntryCount++;
                    continue;
                }

                entries.Add(new WordIndexConcordanceEntry(searchText, indexText));
            }

            return entries;
        }

        private static IEnumerable<Paragraph> EnumerateConcordanceTargetParagraphs(Body body) {
            foreach (OpenXmlElement child in body.ChildElements) {
                if (child is Paragraph paragraph) {
                    yield return paragraph;
                    continue;
                }

                if (child is Table table) {
                    foreach (Paragraph nestedParagraph in table.Descendants<Paragraph>()) {
                        yield return nestedParagraph;
                    }

                    continue;
                }

                if (child is SdtBlock sdtBlock && !IsTableOfContentsBlock(sdtBlock)) {
                    foreach (Paragraph nestedParagraph in sdtBlock.Descendants<Paragraph>()) {
                        yield return nestedParagraph;
                    }
                }
            }
        }

        private static bool IsTableOfContentsBlock(SdtBlock sdtBlock) {
            return sdtBlock.Descendants<DocPartGallery>()
                .Any(gallery => string.Equals(gallery.Val?.Value, "Table of Contents", StringComparison.OrdinalIgnoreCase));
        }

        private static bool ContainsConcordanceMatch(string text, string searchText, bool matchCase, bool matchWholeWord) {
            StringComparison comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            int startIndex = 0;
            while (startIndex <= text.Length) {
                int index = text.IndexOf(searchText, startIndex, comparison);
                if (index < 0) {
                    return false;
                }

                if (!matchWholeWord || HasWordBoundary(text, index, searchText.Length)) {
                    return true;
                }

                startIndex = index + searchText.Length;
            }

            return false;
        }

        private static bool HasWordBoundary(string text, int index, int length) {
            bool leftBoundary = index == 0 || !char.IsLetterOrDigit(text[index - 1]);
            int rightIndex = index + length;
            bool rightBoundary = rightIndex >= text.Length || !char.IsLetterOrDigit(text[rightIndex]);
            return leftBoundary && rightBoundary;
        }

        private static bool CanUseIndexConcordanceText(string value) {
            return value.Length > 0 &&
                value.IndexOf('"') < 0 &&
                value.IndexOf('\\') < 0 &&
                !value.Any(char.IsControl);
        }

        private static string CreateIndexConcordanceInstruction(string indexText) {
            return " XE \"" + indexText + "\" ";
        }

        private static SimpleField CreateHiddenIndexConcordanceField(string instruction) {
            return new SimpleField(
                new Run(
                    new RunProperties(
                        new Vanish()),
                    new Text(string.Empty))) {
                Instruction = instruction
            };
        }

        private static bool HasIndexConcordanceInstruction(Paragraph paragraph, string instruction) {
            string normalized = NormalizeIndexConcordanceInstruction(instruction);
            foreach (SimpleField field in paragraph.Descendants<SimpleField>()) {
                if (NormalizeIndexConcordanceInstruction(field.Instruction?.Value ?? string.Empty) == normalized) {
                    return true;
                }
            }

            foreach (FieldCode fieldCode in paragraph.Descendants<FieldCode>()) {
                if (NormalizeIndexConcordanceInstruction(fieldCode.Text ?? string.Empty) == normalized) {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizeIndexConcordanceInstruction(string instruction) {
            return string.Join(" ", instruction.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
        }

        private static string GetParagraphText(Paragraph paragraph) {
            return GetElementText(paragraph);
        }

        private static string GetElementText(OpenXmlElement element) {
            return string.Concat(element.Descendants<Text>().Select(text => text.Text));
        }
    }
}
