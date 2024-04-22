using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordSection {

        /// <summary>
        /// Converts tables to WordTables
        /// </summary>
        /// <param name="document"></param>
        /// <param name="tables"></param>
        /// <returns></returns>
        internal static List<WordTable> ConvertTableToWordTable(WordDocument document, IEnumerable<Table> tables) {
            var list = new List<WordTable>();
            foreach (Table table in tables) {
                list.Add(new WordTable(document, table));
            }
            return list;
        }

        /// <summary>
        /// Converts SdtBlock to WordWatermark if it's a watermark
        /// Hopefully detection is good enough, but it's possible that it will catch other things as well
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlock"></param>
        /// <returns></returns>
        internal static List<WordWatermark> ConvertStdBlockToWatermark(WordDocument document, IEnumerable<SdtBlock> sdtBlock) {
            var list = new List<WordWatermark>();
            foreach (SdtBlock block in sdtBlock) {
                var sdtContent = block.SdtContentBlock;
                if (sdtContent == null) {
                    continue;
                }
                var paragraphs = sdtContent.ChildElements.OfType<Paragraph>().FirstOrDefault();
                if (paragraphs == null) {
                    continue;
                }
                var run = paragraphs.ChildElements.OfType<Run>().FirstOrDefault();
                if (run == null) {
                    continue;
                }
                var picture = run.ChildElements.OfType<Picture>().FirstOrDefault();
                if (picture == null) {
                    continue;
                }
                list.Add(new WordWatermark(document, block));
            }
            return list;
        }

        /// <summary>
        /// Converts paragraphs to WordParagraphs
        /// </summary>
        /// <param name="document"></param>
        /// <param name="paragraphs"></param>
        /// <returns></returns>
        internal static List<WordParagraph> ConvertParagraphsToWordParagraphs(WordDocument document, IEnumerable<Paragraph> paragraphs) {
            var list = new List<WordParagraph>();
            Dictionary<BookmarkStart, WordBookmark> bookmarks = new Dictionary<BookmarkStart, WordBookmark>();

            foreach (Paragraph paragraph in paragraphs) {
                var childElements = paragraph.ChildElements;
                if (childElements.Count == 1 && childElements[0] is ParagraphProperties) {
                    // basically empty, we still want to track it, but that's about it
                    list.Add(new WordParagraph(document, paragraph));
                } else if (childElements.Any()) {
                    List<Run> runList = new List<Run>();
                    bool foundField = false;
                    foreach (var element in paragraph.ChildElements) {
                        WordParagraph wordParagraph;
                        if (element is Run) {
                            var run = (Run)element;
                            var fieldChar = run.ChildElements.OfType<FieldChar>().FirstOrDefault();
                            if (foundField == true) {
                                if (fieldChar != null && fieldChar.FieldCharType == FieldCharValues.End) {
                                    foundField = false;
                                    runList.Add(run);
                                    wordParagraph = new WordParagraph(document, paragraph, runList);
                                    list.Add(wordParagraph);
                                } else {
                                    runList.Add(run);
                                }
                            } else {
                                if (fieldChar != null) {
                                    if (fieldChar.FieldCharType == FieldCharValues.Begin) {
                                        runList.Add(run);
                                        foundField = true;
                                    }
                                } else {
                                    wordParagraph = new WordParagraph(document, paragraph, run);
                                    list.Add(wordParagraph);
                                }
                            }
                        } else if (element is Hyperlink) {
                            wordParagraph = new WordParagraph(document, paragraph, (Hyperlink)element);
                            list.Add(wordParagraph);
                        } else if (element is SimpleField) {
                            wordParagraph = new WordParagraph(document, paragraph, (SimpleField)element);
                            list.Add(wordParagraph);
                        } else if (element is BookmarkStart) {
                            wordParagraph = new WordParagraph(document, paragraph, (BookmarkStart)element);
                            list.Add(wordParagraph);
                        } else if (element is BookmarkEnd) {
                            // not needed, we will search for matching bookmark end in the bookmark (i guess)
                        } else if (element is DocumentFormat.OpenXml.Math.OfficeMath) {
                            wordParagraph = new WordParagraph(document, paragraph, (DocumentFormat.OpenXml.Math.OfficeMath)element);
                            list.Add(wordParagraph);
                        } else if (element is DocumentFormat.OpenXml.Math.Paragraph) {
                            wordParagraph = new WordParagraph(document, paragraph, (DocumentFormat.OpenXml.Math.Paragraph)element);
                            list.Add(wordParagraph);
                        } else if (element is SdtRun) {
                            list.Add(new WordParagraph(document, paragraph, (SdtRun)element));
                        } else if (element is ProofError) {

                        } else if (element is ParagraphProperties) {

                        } else {
                            Debug.WriteLine("Please implement me! " + element.GetType().Name);
                        }
                    }
                } else {
                    // add empty word paragraph
                    list.Add(new WordParagraph(document, paragraph));
                }
            }

            return list;
        }

        /// <summary>
        /// Get all paragraphs in given section
        /// </summary>
        /// <returns></returns>
        private List<WordParagraph> GetParagraphsList() {
            Dictionary<int, List<Paragraph>> dataSections = new Dictionary<int, List<Paragraph>>();
            var count = 0;

            dataSections[count] = new List<Paragraph>();
            var foundCount = -1;
            if (_wordprocessingDocument.MainDocumentPart.Document != null) {
                if (_wordprocessingDocument.MainDocumentPart.Document.Body != null) {
                    var listElements = _wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements;
                    foreach (var element in listElements) {
                        if (element is Paragraph) {
                            Paragraph paragraph = (Paragraph)element;
                            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.SectionProperties != null) {
                                if (paragraph.ParagraphProperties.SectionProperties == _sectionProperties) {
                                    foundCount = count;
                                }

                                count++;
                                dataSections[count] = new List<Paragraph>();
                            } else {
                                dataSections[count].Add(paragraph);
                            }
                        }
                    }

                    if (foundCount < 0) {
                        var sectionProperties = _wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                        if (sectionProperties == _sectionProperties) {
                            foundCount = count;
                        }
                    }
                }
            }

            return ConvertParagraphsToWordParagraphs(_document, dataSections[foundCount]);
        }


        /// <summary>
        /// This method gets all lists for given document (for all sections)
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        internal static List<WordList> GetAllDocumentsLists(WordDocument document) {
            var numbering = document._wordprocessingDocument.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) {
                return new List<WordList>(0);
            }

            return numbering.ChildElements.OfType<NumberingInstance>()
                .Select(element => new WordList(document, element.NumberID))
                .ToList();
        }

        /// <summary>
        /// This method gets lists for given section. It's possible that given list spreads thru multiple sections.
        /// Therefore sum of all sections lists doesn't necessary match all lists count for a document.
        /// </summary>
        /// <returns></returns>
        private List<WordList> GetLists() {
            List<WordList> allLists = GetAllDocumentsLists(_document);

            List<WordList> lists = new List<WordList>();
            var usedNumbers = this.ParagraphListItemsNumbers;
            foreach (var list in allLists) {
                if (usedNumbers.Contains(list._numberId)) {
                    lists.Add(list);
                }
            }
            return lists;
        }

        /// <summary>
        /// Gets list of tables in given section
        /// </summary>
        /// <returns></returns>
        private List<WordTable> GetTablesList() {
            Dictionary<int, List<WordTable>> dataSections = new Dictionary<int, List<WordTable>>();
            var count = 0;

            dataSections[count] = new List<WordTable>();
            var foundCount = -1;
            if (_wordprocessingDocument.MainDocumentPart.Document != null) {
                if (_wordprocessingDocument.MainDocumentPart.Document.Body != null) {
                    var listElements = _wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements;
                    foreach (var element in listElements) {
                        if (element is Paragraph) {
                            Paragraph paragraph = (Paragraph)element;
                            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.SectionProperties != null) {
                                if (paragraph.ParagraphProperties.SectionProperties == _sectionProperties) {
                                    foundCount = count;
                                }

                                count++;
                                dataSections[count] = new List<WordTable>();
                            }
                        } else if (element is Table) {
                            WordTable wordTable = new WordTable(_document, (Table)element);
                            dataSections[count].Add(wordTable);
                        }
                    }

                    var sectionProperties = _wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                    if (sectionProperties == _sectionProperties) {
                        foundCount = count;
                    }
                }
            }

            return dataSections[foundCount];
        }

        /// <summary>
        /// Gets list of embedded documents in given section
        /// </summary>
        /// <returns></returns>
        private List<WordEmbeddedDocument> GetEmbeddedDocumentsList() {
            Dictionary<int, List<WordEmbeddedDocument>> dataSections = new Dictionary<int, List<WordEmbeddedDocument>>();
            var count = 0;

            dataSections[count] = new List<WordEmbeddedDocument>();
            var foundCount = -1;
            if (_wordprocessingDocument.MainDocumentPart.Document != null) {
                if (_wordprocessingDocument.MainDocumentPart.Document.Body != null) {
                    var listElements = _wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements;
                    foreach (var element in listElements) {
                        if (element is Paragraph) {
                            Paragraph paragraph = (Paragraph)element;
                            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.SectionProperties != null) {
                                if (paragraph.ParagraphProperties.SectionProperties == _sectionProperties) {
                                    foundCount = count;
                                }

                                count++;
                                dataSections[count] = new List<WordEmbeddedDocument>();
                            }
                        } else if (element is AltChunk) {
                            WordEmbeddedDocument wordEmbeddedDocument = new WordEmbeddedDocument(_document, (AltChunk)element);
                            dataSections[count].Add(wordEmbeddedDocument);
                        }
                    }

                    var sectionProperties = _wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                    if (sectionProperties == _sectionProperties) {
                        foundCount = count;
                    }
                }
            }

            return dataSections[foundCount];
        }

        /// <summary>
        /// Gets list of watermarks in given section
        /// </summary>
        /// <returns></returns>
        private List<SdtBlock> GetSdtBlockList() {
            Dictionary<int, List<SdtBlock>> dataSections = new Dictionary<int, List<SdtBlock>>();
            var count = 0;

            dataSections[count] = new List<SdtBlock>();
            var foundCount = -1;
            if (_wordprocessingDocument.MainDocumentPart.Document != null) {
                if (_wordprocessingDocument.MainDocumentPart.Document.Body != null) {
                    var listElements = _wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements;
                    foreach (var element in listElements) {
                        if (element is Paragraph) {
                            Paragraph paragraph = (Paragraph)element;
                            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.SectionProperties != null) {
                                if (paragraph.ParagraphProperties.SectionProperties == _sectionProperties) {
                                    foundCount = count;
                                }

                                count++;
                                dataSections[count] = new List<SdtBlock>();
                            }
                        } else if (element is SdtBlock) {
                            dataSections[count].Add((SdtBlock)element);
                        }
                    }

                    var sectionProperties = _wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                    if (sectionProperties == _sectionProperties) {
                        foundCount = count;
                    }
                }
            }
            return dataSections[foundCount];
        }

        /// <summary>
        /// This method moves headers and footers and title page to section before it.
        /// It also copies copies all other parts of sections (PageSize,PageMargin and others) to section before it.
        /// This is because headers/footers when applied to section apply to the rest of the document
        /// unless there are headers/footers on next section.
        /// On the other hand page size doesn't apply to other sections
        /// and word uses default values. 
        /// </summary>
        /// <param name="sectionProperties"></param>
        /// <param name="newSectionProperties"></param>
        private static void CopySectionProperties(SectionProperties sectionProperties, SectionProperties newSectionProperties) {
            if (newSectionProperties.ChildElements.Count == 0) {
                var listSectionEntries = sectionProperties.ChildElements.ToList();
                foreach (var element in listSectionEntries) {
                    if (element is HeaderReference) {
                        newSectionProperties.Append(element.CloneNode(true));
                        sectionProperties.RemoveChild(element);
                    } else if (element is FooterReference) {
                        newSectionProperties.Append(element.CloneNode(true));
                        sectionProperties.RemoveChild(element);
                        //} else if (element is PageSize) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                    } else if (element is PageMargin) {
                        newSectionProperties.Append(element.CloneNode(true));
                        //sectionProperties.RemoveChild(element);
                        //} else if (element is Columns) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                        //} else if (element is DocGrid) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                        //} else if (element is SectionType) {
                        //    newSectionProperties.Append(element.CloneNode(true));
                    } else if (element is TitlePage) {
                        newSectionProperties.Append(element.CloneNode(true));
                        sectionProperties.RemoveChild(element);
                    } else {
                        newSectionProperties.Append(element.CloneNode(true));
                        //throw new NotImplementedException("This isn't implemented yet?");
                    }
                }
                //sectionProperties.RemoveAllChildren();
                //newSectionProperties.Append(listSectionEntries);
            }
        }
    }
}
