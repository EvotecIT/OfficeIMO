using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        private static object? ResolveParent(WordDocument document, Paragraph paragraph) {
            if (document == null || paragraph == null) {
                return null;
            }

            var tableCell = paragraph.Ancestors<TableCell>().FirstOrDefault();
            if (tableCell != null) {
                return CreateWordTableCell(document, tableCell);
            }

            var topLevel = GetTopLevelContainer(paragraph);
            return topLevel switch {
                Header header => FindHeader(document, header),
                Footer footer => FindFooter(document, footer),
                Body => FindSection(document, paragraph),
                _ => null
            };
        }

        private static WordTableCell? CreateWordTableCell(WordDocument document, TableCell tableCell) {
            var row = tableCell.Ancestors<TableRow>().FirstOrDefault();
            var table = tableCell.Ancestors<Table>().FirstOrDefault();

            if (row == null || table == null) {
                return null;
            }

            var wordTable = new WordTable(document, table, initializeChildren: false);
            var wordRow = new WordTableRow(wordTable, row, document, initializeCells: false);
            return new WordTableCell(document, wordTable, wordRow, tableCell, ensureCellProperties: false);
        }

        private static WordSection? FindSection(WordDocument document, Paragraph paragraph) {
            var sectionProps = GetSectionPropertiesForElement(paragraph);
            if (sectionProps != null) {
                foreach (var section in document.Sections) {
                    if (ReferenceEquals(section._sectionProperties, sectionProps)) {
                        return section;
                    }

                    if (AreSectionsEquivalent(section._sectionProperties, sectionProps)) {
                        return section;
                    }
                }
            }

            if (document.Sections.Count == 1) {
                return document.Sections[0];
            }

            return null;
        }

        private static WordHeader? FindHeader(WordDocument document, Header header) {
            foreach (var section in document.Sections) {
                if (section.Header.Default != null && ReferenceEquals(section.Header.Default._header, header)) {
                    return section.Header.Default;
                }

                if (section.Header.Even != null && ReferenceEquals(section.Header.Even._header, header)) {
                    return section.Header.Even;
                }

                if (section.Header.First != null && ReferenceEquals(section.Header.First._header, header)) {
                    return section.Header.First;
                }
            }

            return null;
        }

        private static WordFooter? FindFooter(WordDocument document, Footer footer) {
            foreach (var section in document.Sections) {
                if (section.Footer.Default != null && ReferenceEquals(section.Footer.Default._footer, footer)) {
                    return section.Footer.Default;
                }

                if (section.Footer.Even != null && ReferenceEquals(section.Footer.Even._footer, footer)) {
                    return section.Footer.Even;
                }

                if (section.Footer.First != null && ReferenceEquals(section.Footer.First._footer, footer)) {
                    return section.Footer.First;
                }
            }

            return null;
        }
    }
}
