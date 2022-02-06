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
        internal static List<WordParagraph> ConvertParagraphsToWordParagraphs(WordDocument document, IEnumerable<Paragraph> paragraphs) {
            var list = new List<WordParagraph>();
            foreach (Paragraph paragraph in paragraphs) {
                //WordParagraph wordParagraph = new WordParagraph(_document, paragraph, null);

                int count = 0;
                var listRuns = paragraph.ChildElements.OfType<Run>();
                if (listRuns.Any()) {
                    foreach (var run in paragraph.ChildElements.OfType<Run>()) {
                        RunProperties runProperties = run.RunProperties;
                        Text text = run.ChildElements.OfType<Text>().FirstOrDefault();
                        Drawing drawing = run.ChildElements.OfType<Drawing>().FirstOrDefault();

                        WordImage newImage = null;
                        if (drawing != null) {
                            newImage = new WordImage(document, drawing);
                        }

                        WordParagraph wordParagraph;
                        if (count > 0) {
                            wordParagraph = new WordParagraph(document, false, paragraph, paragraph.ParagraphProperties, runProperties, run);
                            wordParagraph.Image = newImage;

                            if (wordParagraph.IsPageBreak) {
                                // document._currentSection.PageBreaks.Add(wordParagraph);
                            }

                            if (wordParagraph.IsListItem) {
                                //LoadListToDocument(document, wordParagraph);
                            }

                            list.Add(wordParagraph);
                        } else {
                            wordParagraph = new WordParagraph(document, false, paragraph, paragraph.ParagraphProperties, runProperties, run);

                            if (newImage != null) {
                                wordParagraph.Image = newImage;
                            }

                            if (wordParagraph.IsPageBreak) {
                                // section.PageBreaks.Add(this);
                            }

                            if (wordParagraph.IsListItem) {
                                //LoadListToDocument(document, this);
                            }

                            list.Add(wordParagraph);
                        }

                        count++;
                    }
                } else {
                    // add empty word paragraph
                    list.Add(new WordParagraph(document, false, paragraph, null, null, null));
                }
            }

            return list;
        }

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
                            WordTable wordTable = new WordTable(_document, null, (Table)element);
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