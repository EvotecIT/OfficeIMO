using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordTableCell {
        internal TableCell _tableCell;
        internal TableCellProperties _tableCellProperties;

        private List<WordParagraph> GetParagraphs(IEnumerable<Paragraph> paragraphs) {
            var list = new List<WordParagraph>();
            foreach (Paragraph paragraph in paragraphs) {
                WordParagraph wordParagraph = new WordParagraph(_document, paragraph, null);

                int count = 0;
                var listRuns = paragraph.ChildElements.OfType<Run>();
                if (listRuns.Any()) {
                    foreach (var run in paragraph.ChildElements.OfType<Run>()) {
                        RunProperties runProperties = run.RunProperties;
                        Text text = run.ChildElements.OfType<Text>().FirstOrDefault();
                        Drawing drawing = run.ChildElements.OfType<Drawing>().FirstOrDefault();

                        WordImage newImage = null;
                        if (drawing != null) {
                            newImage = new WordImage(_document, drawing);
                        }

                        if (count > 0) {
                            wordParagraph = new WordParagraph(_document);
                            wordParagraph._document = _document;
                            wordParagraph._run = run;
                            wordParagraph._text = text;
                            wordParagraph._paragraph = paragraph;
                            wordParagraph._paragraphProperties = paragraph.ParagraphProperties;
                            wordParagraph._runProperties = runProperties;
                            //wordParagraph._section = section;

                            wordParagraph.Image = newImage;

                            if (wordParagraph.IsPageBreak) {
                                // document._currentSection.PageBreaks.Add(wordParagraph);
                            }

                            if (wordParagraph.IsListItem) {
                                //LoadListToDocument(document, wordParagraph);
                            }

                            list.Add(wordParagraph);
                        } else {
                            // wordParagraph._document = document;
                            wordParagraph._run = run;
                            wordParagraph._text = text;
                            wordParagraph._paragraph = paragraph;
                            wordParagraph._paragraphProperties = paragraph.ParagraphProperties;
                            wordParagraph._runProperties = runProperties;
                            // wordParagraph._section = section;

                            if (newImage != null) {
                                wordParagraph.Image = newImage;
                            }

                            // this is to prevent adding Tables Paragraphs to section Paragraphs
                            //if (section != null) {
                            // section.Paragraphs.Add(this);
                            if (wordParagraph.IsPageBreak) {
                                // section.PageBreaks.Add(this);
                            }
                            //}

                            if (wordParagraph.IsListItem) {
                                //LoadListToDocument(document, this);
                            }
                            list.Add(wordParagraph);
                        }
                        count++;
                    }
                } else {
                    // add empty word paragraph
                    list.Add(wordParagraph);
                }
            }

            return list;
        }

        public List<WordParagraph> Paragraphs => GetParagraphs(_tableCell.ChildElements.OfType<Paragraph>());

        private readonly WordTable _wordTable;
        private readonly WordTableRow _wordTableRow;
        private readonly WordDocument _document;

        public WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow) {
            TableCell tableCell = new TableCell();
            TableCellProperties tableCellProperties = new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" });

            // Specify the width property of the table cell.
            tableCell.Append(tableCellProperties);

            // Specify the table cell content.
            //tableCell.Append(new Paragraph(new Run(new Text("Hello, World!"))));

            WordParagraph paragraph = new WordParagraph();
            //tableCell.Append(new Paragraph(new Run(new Text("Hello, World!"))));
            //Paragraphs.Add(paragraph);

            tableCell.Append(paragraph._paragraph);

            wordTableRow._tableRow.Append(tableCell);

            _tableCellProperties = tableCellProperties;
            _tableCell = tableCell;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _document = document;
        }

        public WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow, TableCell tableCell) {
            _tableCell = tableCell;
            _tableCellProperties = tableCell.TableCellProperties;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _document = document;

            //foreach (Paragraph paragraph in tableCell.ChildElements.OfType<Paragraph>().ToList()) {
            //    WordParagraph wordParagraph = new WordParagraph(document, paragraph, null);
            //    this.Paragraphs.Add(wordParagraph);
            //}
        }

        public void Remove() {
            _tableCell.Remove();
            //_wordTableRow.Cells.Remove(this);

        }
    }
}
