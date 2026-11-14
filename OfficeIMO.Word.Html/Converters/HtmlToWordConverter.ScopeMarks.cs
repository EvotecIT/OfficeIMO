using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        /// <summary>
        /// Remembers where a scope ends before a container adds content to it. Body scopes keep the last
        /// paragraph and block, so generated content is found by walking back from the end of the section
        /// instead of listing the whole document for every container. Cell and header scopes are small and
        /// keep positions in their paragraph and table lists.
        /// </summary>
        private readonly struct ScopeMark {
            private ScopeMark(bool isBodyMark, OpenXmlElement? lastBlock, Paragraph? lastParagraph, int lastParagraphWrapperCount, int paragraphStartIndex, int tableStartIndex) {
                IsBodyMark = isBodyMark;
                LastBlock = lastBlock;
                LastParagraph = lastParagraph;
                LastParagraphWrapperCount = lastParagraphWrapperCount;
                ParagraphStartIndex = paragraphStartIndex;
                TableStartIndex = tableStartIndex;
            }

            internal static ScopeMark ForBody(OpenXmlElement? lastBlock, Paragraph? lastParagraph, int lastParagraphWrapperCount) =>
                new(true, lastBlock, lastParagraph, lastParagraphWrapperCount, 0, 0);

            internal static ScopeMark ForList(int paragraphStartIndex, int tableStartIndex) =>
                new(false, null, null, 0, paragraphStartIndex, tableStartIndex);

            internal bool IsBodyMark { get; }
            internal OpenXmlElement? LastBlock { get; }
            internal Paragraph? LastParagraph { get; }
            internal int LastParagraphWrapperCount { get; }
            internal int ParagraphStartIndex { get; }
            internal int TableStartIndex { get; }
        }

        /// <param name="section">Section that receives body content.</param>
        /// <param name="cell">Cell that receives content, if any.</param>
        /// <param name="headerFooter">Header or footer that receives content, if any.</param>
        /// <param name="replaceableCellParagraphIsGenerated">Counts a lone empty cell paragraph as generated, because the next added paragraph replaces it.</param>
        /// <param name="includeTables">Records the table position as well; only callers that read generated tables need it.</param>
        private static ScopeMark MarkScope(
            WordSection section,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter,
            bool replaceableCellParagraphIsGenerated = true,
            bool includeTables = true) {
            if (cell == null && headerFooter == null &&
                section.TryGetLastBlocks(out OpenXmlElement? lastBlock, out Paragraph? lastParagraph)) {
                int wrapperCount = lastParagraph == null ? 0 : ProjectParagraph(section, lastParagraph).Count;
                return ScopeMark.ForBody(lastBlock, lastParagraph, wrapperCount);
            }

            List<WordParagraph> paragraphs = GetParagraphsInScope(section, cell, headerFooter);
            int paragraphStartIndex = replaceableCellParagraphIsGenerated &&
                                      cell != null &&
                                      paragraphs.Count == 1 &&
                                      IsReplaceableEmptyCellParagraph(paragraphs[0])
                ? 0
                : paragraphs.Count;
            int tableStartIndex = includeTables ? GetTablesInScope(section, cell, headerFooter).Count : 0;
            return ScopeMark.ForList(paragraphStartIndex, tableStartIndex);
        }

        private static List<WordParagraph> ProjectParagraph(WordSection section, Paragraph paragraph) =>
            WordSection.ConvertParagraphToWordParagraphs(section._document, paragraph);

        private static bool TryGetBodyParagraphsAfter(WordSection section, ScopeMark mark, List<Paragraph> paragraphs) =>
            section.TryGetBlocksAfter(mark.LastBlock, mark.LastParagraph, paragraphs, new List<Table>());

        private static List<WordParagraph> GetGeneratedParagraphs(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter, ScopeMark mark) {
            if (!mark.IsBodyMark) {
                return GetParagraphsInScope(section, cell, headerFooter).Skip(mark.ParagraphStartIndex).ToList();
            }

            var paragraphs = new List<Paragraph>();
            if (!TryGetBodyParagraphsAfter(section, mark, paragraphs)) {
                return GetSectionParagraphsAfter(section, mark);
            }

            var generated = new List<WordParagraph>();
            if (mark.LastParagraph != null) {
                generated.AddRange(ProjectParagraph(section, mark.LastParagraph).Skip(mark.LastParagraphWrapperCount));
            }
            foreach (Paragraph paragraph in paragraphs) {
                generated.AddRange(ProjectParagraph(section, paragraph));
            }
            return generated;
        }

        /// <summary>Returns the first paragraph added to the scope after <paramref name="mark"/>, if any.</summary>
        private static WordParagraph? GetFirstGeneratedParagraph(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter, ScopeMark mark) {
            if (!mark.IsBodyMark) {
                List<WordParagraph> scoped = GetParagraphsInScope(section, cell, headerFooter);
                return scoped.Count > mark.ParagraphStartIndex ? scoped[mark.ParagraphStartIndex] : null;
            }

            var paragraphs = new List<Paragraph>();
            if (!TryGetBodyParagraphsAfter(section, mark, paragraphs)) {
                return GetSectionParagraphsAfter(section, mark).FirstOrDefault();
            }

            if (mark.LastParagraph != null) {
                List<WordParagraph> projected = ProjectParagraph(section, mark.LastParagraph);
                if (projected.Count > mark.LastParagraphWrapperCount) {
                    return projected[mark.LastParagraphWrapperCount];
                }
            }
            foreach (Paragraph paragraph in paragraphs) {
                List<WordParagraph> projected = ProjectParagraph(section, paragraph);
                if (projected.Count > 0) {
                    return projected[0];
                }
            }
            return null;
        }

        /// <summary>Returns the last paragraph of the scope when content was added to it after <paramref name="mark"/>.</summary>
        private static WordParagraph? GetLastGeneratedParagraph(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter, ScopeMark mark) {
            if (!mark.IsBodyMark) {
                List<WordParagraph> scoped = GetParagraphsInScope(section, cell, headerFooter);
                return scoped.Count > mark.ParagraphStartIndex ? scoped[scoped.Count - 1] : null;
            }

            var paragraphs = new List<Paragraph>();
            if (!TryGetBodyParagraphsAfter(section, mark, paragraphs)) {
                return GetSectionParagraphsAfter(section, mark).LastOrDefault();
            }

            for (int index = paragraphs.Count - 1; index >= 0; index--) {
                List<WordParagraph> projected = ProjectParagraph(section, paragraphs[index]);
                if (projected.Count > 0) {
                    return projected[projected.Count - 1];
                }
            }
            if (mark.LastParagraph != null) {
                List<WordParagraph> projected = ProjectParagraph(section, mark.LastParagraph);
                if (projected.Count > mark.LastParagraphWrapperCount) {
                    return projected[projected.Count - 1];
                }
            }
            return null;
        }

        /// <summary>Full-list fallback for body marks when the section layout cannot be walked from its end.</summary>
        private static List<WordParagraph> GetSectionParagraphsAfter(WordSection section, ScopeMark mark) {
            List<WordParagraph> paragraphs = section.Paragraphs;
            if (mark.LastParagraph == null) {
                return paragraphs;
            }

            int anchorIndex = paragraphs.FindIndex(paragraph => ReferenceEquals(paragraph._paragraph, mark.LastParagraph));
            return anchorIndex < 0
                ? new List<WordParagraph>()
                : paragraphs.Skip(anchorIndex + mark.LastParagraphWrapperCount).ToList();
        }

        private static List<WordTable> GetGeneratedTables(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter, ScopeMark mark) {
            if (!mark.IsBodyMark) {
                return GetTablesInScope(section, cell, headerFooter).Skip(mark.TableStartIndex).ToList();
            }

            var tables = new List<Table>();
            if (!section.TryGetBlocksAfter(mark.LastBlock, mark.LastParagraph, new List<Paragraph>(), tables)) {
                return GetSectionTablesAfter(section, mark.LastBlock);
            }

            var generated = new List<WordTable>(tables.Count);
            foreach (Table table in tables) {
                generated.Add(new WordTable(section._document, table));
            }
            return generated;
        }

        /// <summary>Full-list fallback for body marks when the section layout cannot be walked from its end.</summary>
        private static List<WordTable> GetSectionTablesAfter(WordSection section, OpenXmlElement? lastBlock) {
            List<WordTable> tables = section.Tables;
            if (lastBlock == null) {
                return tables;
            }

            var following = new HashSet<OpenXmlElement>();
            if (lastBlock.Parent != null) {
                for (OpenXmlElement? sibling = lastBlock.NextSibling(); sibling != null; sibling = sibling.NextSibling()) {
                    if (sibling is Table) {
                        following.Add(sibling);
                    }
                }
            }
            return tables.Where(table => following.Contains(table._table)).ToList();
        }

        private static bool IsParagraphInScope(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter, ScopeMark mark, WordParagraph paragraph) {
            if (mark.IsBodyMark &&
                mark.LastParagraph != null &&
                ReferenceEquals(paragraph._paragraph, mark.LastParagraph) &&
                TryGetBodyParagraphsAfter(section, mark, new List<Paragraph>())) {
                return ProjectParagraph(section, mark.LastParagraph).Count > 0;
            }

            return GetParagraphsInScope(section, cell, headerFooter)
                .Any(candidate => ReferenceEquals(candidate._paragraph, paragraph._paragraph));
        }
    }
}
