using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    public class WordTableCell {
        public WordTableCellBorder Borders;

        internal TableCell _tableCell;
        internal TableCellProperties _tableCellProperties;

        //private List<WordParagraph> GetParagraphs(IEnumerable<Paragraph> paragraphs) {
        //    var list = new List<WordParagraph>();
        //    foreach (Paragraph paragraph in paragraphs) {
        //        WordParagraph wordParagraph = new WordParagraph(_document, paragraph);

        //        int count = 0;
        //        var listRuns = paragraph.ChildElements.OfType<Run>();
        //        if (listRuns.Any()) {
        //            foreach (var run in paragraph.ChildElements.OfType<Run>()) {
        //                RunProperties runProperties = run.RunProperties;
        //                Text text = run.ChildElements.OfType<Text>().FirstOrDefault();
        //                Drawing drawing = run.ChildElements.OfType<Drawing>().FirstOrDefault();

        //                WordImage newImage = null;
        //                if (drawing != null) {
        //                    newImage = new WordImage(_document, drawing);
        //                }

        //                if (count > 0) {
        //                    wordParagraph = new WordParagraph(_document);
        //                    wordParagraph._document = _document;
        //                    wordParagraph._run = run;
        //                    wordParagraph._text = text;
        //                    wordParagraph._paragraph = paragraph;
        //                    wordParagraph._paragraphProperties = paragraph.ParagraphProperties;
        //                    wordParagraph._runProperties = runProperties;
        //                    //wordParagraph._section = section;

        //                    wordParagraph.Image = newImage;

        //                    if (wordParagraph.IsPageBreak) {
        //                        // document._currentSection.PageBreaks.Add(wordParagraph);
        //                    }

        //                    if (wordParagraph.IsListItem) {
        //                        //LoadListToDocument(document, wordParagraph);
        //                    }

        //                    list.Add(wordParagraph);
        //                } else {
        //                    // wordParagraph._document = document;
        //                    wordParagraph._run = run;
        //                    wordParagraph._text = text;
        //                    wordParagraph._paragraph = paragraph;
        //                    wordParagraph._paragraphProperties = paragraph.ParagraphProperties;
        //                    wordParagraph._runProperties = runProperties;
        //                    // wordParagraph._section = section;

        //                    if (newImage != null) {
        //                        wordParagraph.Image = newImage;
        //                    }

        //                    // this is to prevent adding Tables Paragraphs to section Paragraphs
        //                    //if (section != null) {
        //                    // section.Paragraphs.Add(this);
        //                    if (wordParagraph.IsPageBreak) {
        //                        // section.PageBreaks.Add(this);
        //                    }
        //                    //}

        //                    if (wordParagraph.IsListItem) {
        //                        //LoadListToDocument(document, this);
        //                    }

        //                    list.Add(wordParagraph);
        //                }

        //                count++;
        //            }
        //        } else {
        //            // add empty word paragraph
        //            list.Add(wordParagraph);
        //        }
        //    }

        //    return list;
        //}

        //public List<WordParagraph> Paragraphs => GetParagraphs(_tableCell.ChildElements.OfType<Paragraph>());
        public List<WordParagraph> Paragraphs => WordSection.ConvertParagraphsToWordParagraphs(_document, _tableCell.ChildElements.OfType<Paragraph>());
        private readonly WordTable _wordTable;
        private readonly WordTableRow _wordTableRow;
        private readonly WordDocument _document;

        /// <summary>
        /// Gets or Sets Horizontal Merge for a Table Cell
        /// </summary>
        public MergedCellValues? HorizontalMerge {
            get {
                if (_tableCellProperties.HorizontalMerge != null) {
                    return _tableCellProperties.HorizontalMerge.Val;
                }
                return null;
            }
            set {
                if (value == null) {
                    _tableCellProperties.HorizontalMerge.Remove();
                } else {
                    if (_tableCellProperties.HorizontalMerge == null) {
                        _tableCellProperties.HorizontalMerge = new HorizontalMerge();
                    }

                    _tableCellProperties.HorizontalMerge.Val = value;
                }
            }
        }

        /// <summary>
        /// Gets or Sets Vertical Merge for a Table Cell
        /// </summary>
        public MergedCellValues? VerticalMerge {
            get {
                if (_tableCellProperties.VerticalMerge != null) {
                    return _tableCellProperties.VerticalMerge.Val;
                }

                return null;
            }
            set {
                if (value == null) {
                    _tableCellProperties.VerticalMerge.Remove();
                } else {
                    if (_tableCellProperties.VerticalMerge == null) {
                        _tableCellProperties.VerticalMerge = new VerticalMerge();
                    }

                    _tableCellProperties.VerticalMerge.Val = value;
                }
            }
        }

        /// <summary>
        /// Get or set the background color of the cell using hexadecimal color code.
        /// </summary>
        public string ShadingFillColorHex {
            get {
                if (_tableCellProperties.Shading != null) {
                    if (_tableCellProperties.Shading.Fill != null) {
                        return _tableCellProperties.Shading.Fill.Value;
                    }
                }
                return "";
            }
            set {
                if (value != "") {
                    var color = value.Replace("#", "");
                    if (_tableCellProperties.Shading == null) {
                        _tableCellProperties.Shading = new Shading();
                    }
                    _tableCellProperties.Shading.Fill = color;
                    if (_tableCellProperties.Shading.Val == null) {
                        _tableCellProperties.Shading.Val = ShadingPatternValues.Clear;
                    }
                } else {
                    if (_tableCellProperties.Shading != null && _tableCellProperties.Shading.Fill != null) {
                        _tableCellProperties.Shading.Remove();
                    }
                }

            }
        }

        /// <summary>
        /// Get or set the background pattern of a cell
        /// </summary>
        public ShadingPatternValues? ShadingPattern {
            get {
                if (_tableCellProperties.Shading != null) {
                    return _tableCellProperties.Shading.Val;
                }

                return null;
            }
            set {
                if (value != null) {
                    if (_tableCellProperties.Shading == null) {
                        _tableCellProperties.Shading = new Shading();
                    }
                    _tableCellProperties.Shading.Val = value;
                } else {
                    if (_tableCellProperties.Shading != null) {
                        _tableCellProperties.Shading.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Get or set the background color of a cell using SixLabors.Color
        /// </summary>
        public Color? ShadingFillColor {
            get {
                if (ShadingFillColorHex != "") {
                    return SixLabors.ImageSharp.Color.Parse("#" + ShadingFillColorHex);
                }

                return null;
            }
            set {
                if (value != null) {
                    this.ShadingFillColorHex = value.Value.ToHexColor();
                }
            }
        }

        /// <summary>
        /// Gets or sets cell width
        /// </summary>
        public int? Width {
            get {
                if (_tableCellProperties.TableCellWidth != null) {
                    return int.Parse(_tableCellProperties.TableCellWidth.Width);
                }

                return null;
            }
            set {
                if (value != null) {
                    if (_tableCellProperties.TableCellWidth == null) {
                        _tableCellProperties.TableCellWidth = new TableCellWidth();
                    }

                    _tableCellProperties.TableCellWidth.Width = value.ToString();
                } else {
                    if (_tableCellProperties.TableCellWidth != null) {
                        _tableCellProperties.TableCellWidth.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets text direction in a Table Cell
        /// </summary>
        public TextDirectionValues? TextDirection {
            get {
                if (_tableCellProperties.TextDirection != null) {
                    return _tableCellProperties.TextDirection.Val;
                }

                return null;
            }
            set {
                if (value != null) {
                    if (_tableCellProperties.TextDirection == null) {
                        _tableCellProperties.TextDirection = new TextDirection();
                    }
                    _tableCellProperties.TextDirection.Val = value;
                } else {
                    if (_tableCellProperties.TextDirection != null) {
                        _tableCellProperties.TextDirection.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Create a WordTableCell and add it to given Table Row
        /// </summary>
        /// <param name="document"></param>
        /// <param name="wordTable"></param>
        /// <param name="wordTableRow"></param>
        public WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow) {
            TableCell tableCell = new TableCell();
            TableCellProperties tableCellProperties = new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" });

            // Specify the width property of the table cell.
            tableCell.Append(tableCellProperties);

            WordParagraph paragraph = new WordParagraph();

            tableCell.Append(paragraph._paragraph);

            wordTableRow._tableRow.Append(tableCell);

            this.Borders = new WordTableCellBorder(document, wordTable, wordTableRow, this);

            _tableCellProperties = tableCellProperties;
            _tableCell = tableCell;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _document = document;
        }

        /// <summary>
        /// Create a WordTableCell and add it to given Table Row from TableCell
        /// Mostly used for loading TableCells during Document Load
        /// </summary>
        /// <param name="document"></param>
        /// <param name="wordTable"></param>
        /// <param name="wordTableRow"></param>
        /// <param name="tableCell"></param>
        internal WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow, TableCell tableCell) {
            _tableCell = tableCell;
            _tableCellProperties = tableCell.TableCellProperties;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _document = document;

            this.Borders = new WordTableCellBorder(document, wordTable, wordTableRow, this);
        }

        /// <summary>
        /// Remove a cell from a table
        /// </summary>
        public void Remove() {
            _tableCell.Remove();
        }

        /// <summary>
        /// Merges two or more cells together horizontally.
        /// Provides ability to move or delete content of merged cells into single cell
        /// </summary>
        /// <param name="cellsCount"></param>
        /// <param name="copyParagraphs"></param>
        public void MergeHorizontally(int cellsCount, bool copyParagraphs = false) {
            var temporaryCell = _tableCell;
            _tableCell.TableCellProperties.HorizontalMerge = new HorizontalMerge {
                Val = MergedCellValues.Restart
            };

            for (int i = 0; i < cellsCount; i++) {
                if (_tableCell != null) {
                    _tableCell = (TableCell)_tableCell.NextSibling();
                    if (_tableCell != null) {
                        if (copyParagraphs) {
                            // lets find all paragraphs and move them to first table cell
                            var paragraphs = _tableCell.ChildElements.OfType<Paragraph>();
                            foreach (var paragraph in paragraphs) {
                                // moving paragraphs
                                paragraph.Remove();
                                temporaryCell.Append(paragraph);
                            }

                            // but tableCell requires at least one empty paragraph so we provide that request
                            _tableCell.Append(new Paragraph());
                        } else {
                            // lets find all paragraphs and delete them
                            var paragraphs = _tableCell.ChildElements.OfType<Paragraph>();
                            foreach (var paragraph in paragraphs) {
                                paragraph.Remove();
                            }

                            // but tableCell requires at least one empty paragraph so we provide that request
                            _tableCell.Append(new Paragraph());
                        }

                        // then for every table cell we need to continue merging until cellsCount
                        _tableCell.TableCellProperties.HorizontalMerge = new HorizontalMerge {
                            Val = MergedCellValues.Continue
                        };
                    }
                }
            }

        }

        /// <summary>
        /// Splits (unmerge) cells that were merged 
        /// </summary>
        /// <param name="cellsCount"></param>
        public void SplitHorizontally(int cellsCount) {
            if (_tableCell.TableCellProperties.HorizontalMerge != null) {
                _tableCell.TableCellProperties.HorizontalMerge.Remove();
            }
            for (int i = 0; i < cellsCount; i++) {
                if (_tableCell != null) {
                    _tableCell = (TableCell)_tableCell.NextSibling();
                    if (_tableCell != null) {
                        if (_tableCell.TableCellProperties.HorizontalMerge != null) {
                            _tableCell.TableCellProperties.HorizontalMerge.Remove();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Merges two or more cells together vertically 
        /// </summary>
        /// <param name="cellsCount"></param>
        /// <param name="copyParagraphs"></param>
        public void MergeVertically(int cellsCount, bool copyParagraphs = false) {
            var temporaryCell = _tableCell;
            _tableCell.TableCellProperties.VerticalMerge = new VerticalMerge {
                Val = MergedCellValues.Restart
            };
            var tableRow = _tableCell.Parent;
            var indexOfCell = tableRow.ChildElements.ToList().IndexOf(_tableCell);

            for (int i = 0; i < cellsCount; i++) {
                if (_tableCell != null) {
                    if (tableRow != null) {
                        tableRow = tableRow.NextSibling();
                        if (tableRow != null) {
                            // we need to find cell with proper index
                            var tableCells = tableRow.ChildElements.OfType<TableCell>().ToList()[indexOfCell];
                            if (tableCells != null) {
                                _tableCell = tableCells;
                                if (_tableCell != null) {
                                    if (copyParagraphs) {
                                        // lets find all paragraphs and move them to first table cell
                                        var paragraphs = _tableCell.ChildElements.OfType<Paragraph>();
                                        foreach (var paragraph in paragraphs) {
                                            // moving paragraphs
                                            paragraph.Remove();
                                            temporaryCell.Append(paragraph);
                                        }

                                        // but tableCell requires at least one empty paragraph so we provide that request
                                        _tableCell.Append(new Paragraph());
                                    } else {
                                        // lets find all paragraphs and delete them
                                        var paragraphs = _tableCell.ChildElements.OfType<Paragraph>();
                                        foreach (var paragraph in paragraphs) {
                                            paragraph.Remove();
                                        }

                                        // but tableCell requires at least one empty paragraph so we provide that request
                                        _tableCell.Append(new Paragraph());
                                    }

                                    // then for every table cell we need to continue merging until cellsCount
                                    _tableCell.TableCellProperties.VerticalMerge = new VerticalMerge {
                                        Val = MergedCellValues.Continue
                                    };
                                }
                            }
                        }
                    }
                }
            }

        }

        public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this._document, _tableCell, rows, columns, tableStyle);
            return wordTable;
        }
    }
}