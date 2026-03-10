using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using Ovml = DocumentFormat.OpenXml.Vml.Office;

namespace OfficeIMO.Word {
    /// <summary>
    /// Section in WordDocument
    /// </summary>
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
                var watermark = ConvertStdBlockToWatermark(document, block);
                if (watermark != null) {
                    list.Add(watermark);
                }
            }
            return list;
        }

        /// <summary>
        /// Converts SdtBlock to WordWatermark if it's a watermark
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlock"></param>
        /// <returns></returns>
        internal static WordWatermark? ConvertStdBlockToWatermark(WordDocument document, SdtBlock? sdtBlock) {
            if (sdtBlock == null) {
                return null;
            }
            var sdtContent = sdtBlock.SdtContentBlock;
            if (sdtContent == null) {
                return null;
            }
            var paragraphs = sdtContent.ChildElements.OfType<Paragraph>().FirstOrDefault();
            if (paragraphs == null) {
                return null;
            }
            var run = paragraphs.ChildElements.OfType<Run>().FirstOrDefault();
            if (run == null) {
                return null;
            }
            var picture = run.ChildElements.OfType<Picture>().FirstOrDefault();
            if (picture == null) {
                return null;
            }
            return new WordWatermark(document, sdtBlock);
        }

        /// <summary>
        /// Converts StdBlock to WordCoverPage
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlocks"></param>
        /// <returns></returns>
        internal static WordCoverPage? ConvertStdBlockToCoverPage(WordDocument document, IEnumerable<SdtBlock?> sdtBlocks) {
            foreach (var sdtBlock in sdtBlocks) {
                if (sdtBlock == null) {
                    continue;
                }
                var sdtProperties = sdtBlock.ChildElements.OfType<SdtProperties>().FirstOrDefault();
                var docPartObject = sdtProperties?.ChildElements.OfType<SdtContentDocPartObject>().FirstOrDefault();
                var docPartGallery = docPartObject?.ChildElements.OfType<DocPartGallery>().FirstOrDefault();

                if (docPartGallery != null && docPartGallery.Val == "Cover Pages") {
                    return new WordCoverPage(document, sdtBlock!);
                }
            }

            return null;
        }

        /// <summary>
        /// Converts StdBlock to WordTableOfContent
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlocks"></param>
        /// <returns></returns>
        internal static WordTableOfContent? ConvertStdBlockToTableOfContent(WordDocument document, IEnumerable<SdtBlock?> sdtBlocks) {
            foreach (var sdtBlock in sdtBlocks) {
                if (sdtBlock == null) {
                    continue;
                }
                var sdtProperties = sdtBlock.ChildElements.OfType<SdtProperties>().FirstOrDefault();
                var docPartObject = sdtProperties?.ChildElements.OfType<SdtContentDocPartObject>().FirstOrDefault();
                var docPartGallery = docPartObject?.ChildElements.OfType<DocPartGallery>().FirstOrDefault();

                if (docPartGallery != null && docPartGallery.Val == "Table of Contents") {
                    return new WordTableOfContent(document, sdtBlock!, queueUpdateOnOpen: false);
                }
            }
            return null;
        }

        /// <summary>
        /// Converts StdBlock to WordElement
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlock"></param>
        /// <returns></returns>
        internal static WordElement ConvertStdBlockToWordElements(WordDocument document, SdtBlock? sdtBlock) {
            if (sdtBlock == null) {
                return new WordStructuredDocumentTag(document, new SdtBlock());
            }

            var sdtProperties = sdtBlock.ChildElements.OfType<SdtProperties>().FirstOrDefault();
            var docPartObject = sdtProperties?.ChildElements.OfType<SdtContentDocPartObject>().FirstOrDefault();
            var docPartGallery = docPartObject?.ChildElements.OfType<DocPartGallery>().FirstOrDefault();

            if (docPartGallery != null && docPartGallery.Val == "Cover Pages") {
                return new WordCoverPage(document, sdtBlock!);
            } else if (docPartGallery != null && docPartGallery.Val == "Table of Contents") {
                return new WordTableOfContent(document, sdtBlock!);
            }

            var watermark = ConvertStdBlockToWatermark(document, sdtBlock);
            if (watermark != null) {
                return watermark;
            }

            return new WordStructuredDocumentTag(document, sdtBlock!);
        }

        /// <summary>
        /// Converts StdBlock to WordElement
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sdtBlocks"></param>
        /// <returns></returns>
        internal static List<WordElement> ConvertStdBlockToWordElements(WordDocument document, IEnumerable<SdtBlock?> sdtBlocks) {
            var list = new List<WordElement>();
            foreach (var sdtBlock in sdtBlocks) {
                var element = ConvertStdBlockToWordElements(document, sdtBlock);
                if (element != null) {
                    list.Add(element);
                }
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

            foreach (Paragraph paragraph in paragraphs) {
                list.AddRange(ConvertParagraphToWordParagraphs(document, paragraph));
            }

            return list;
        }

        /// <summary>
        /// Converts paragraph to WordParagraphs
        /// </summary>
        /// <param name="document"></param>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        internal static List<WordParagraph> ConvertParagraphToWordParagraphs(WordDocument document, Paragraph paragraph) {
            var list = new List<WordParagraph>();
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
                            if (fieldChar?.FieldCharType?.Value == FieldCharValues.End) {
                                foundField = false;
                                runList.Add(run);
                                wordParagraph = new WordParagraph(document, paragraph, runList);
                                list.Add(wordParagraph);
                                runList = new List<Run>();
                            } else {
                                runList.Add(run);
                            }
                        } else {
                            if (fieldChar?.FieldCharType?.Value == FieldCharValues.Begin) {
                                runList.Add(run);
                                foundField = true;
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
            return list;
        }

        private int GetSectionOrdinal() {
            int sectionIndex = _document.Sections.IndexOf(this);
            if (sectionIndex < 0) {
                throw new InvalidOperationException("The section is not attached to the document.");
            }

            return sectionIndex;
        }

        private int GetSectionCount() {
            return Math.Max(_document.Sections.Count, 1);
        }

        private static bool IsSectionBoundaryParagraph(Paragraph paragraph) {
            return paragraph.ParagraphProperties?.SectionProperties != null;
        }

        private static bool IsPureSectionBreakParagraph(Paragraph paragraph) {
            if (!IsSectionBoundaryParagraph(paragraph)) {
                return false;
            }

            if (paragraph.ChildElements.Any(element => element is not ParagraphProperties)) {
                return false;
            }

            return paragraph.ParagraphProperties?.ChildElements.All(element => element is SectionProperties) != false;
        }

        /// <summary>
        /// Get all paragraphs in given section
        /// </summary>
        /// <returns></returns>
        private List<WordParagraph> GetParagraphsList() {
            int targetSection = GetSectionOrdinal();
            var paragraphsBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<Paragraph>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordParagraph>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is not Paragraph paragraph) {
                    continue;
                }

                if (!IsPureSectionBreakParagraph(paragraph)) {
                    paragraphsBySection[currentSection].Add(paragraph);
                }

                if (IsSectionBoundaryParagraph(paragraph) && currentSection < paragraphsBySection.Count - 1) {
                    currentSection++;
                }
            }

            return ConvertParagraphsToWordParagraphs(_document, paragraphsBySection[targetSection]);
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
                .Select(element => new WordList(document, element.NumberID!.Value))
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
            int targetSection = GetSectionOrdinal();
            var tablesBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<WordTable>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordTable>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is Paragraph paragraph) {
                    if (IsSectionBoundaryParagraph(paragraph) && currentSection < tablesBySection.Count - 1) {
                        currentSection++;
                    }
                } else if (element is Table table) {
                    tablesBySection[currentSection].Add(new WordTable(_document, table));
                }
            }

            return tablesBySection[targetSection];
        }

        /// <summary>
        /// Gets list of embedded documents in given section
        /// </summary>
        /// <returns></returns>
        private List<WordEmbeddedDocument> GetEmbeddedDocumentsList() {
            int targetSection = GetSectionOrdinal();
            var embeddedDocumentsBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<WordEmbeddedDocument>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordEmbeddedDocument>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is Paragraph paragraph) {
                    if (IsSectionBoundaryParagraph(paragraph) && currentSection < embeddedDocumentsBySection.Count - 1) {
                        currentSection++;
                    }
                } else if (element is AltChunk altChunk) {
                    embeddedDocumentsBySection[currentSection].Add(new WordEmbeddedDocument(_document, altChunk));
                }
            }

            return embeddedDocumentsBySection[targetSection];
        }

        private List<WordEmbeddedObject> GetEmbeddedObjectsList() {
            int targetSection = GetSectionOrdinal();
            var embeddedObjectsBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<WordEmbeddedObject>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordEmbeddedObject>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is not Paragraph paragraph) {
                    continue;
                }

                foreach (var run in paragraph.ChildElements.OfType<Run>()) {
                    if (run.Descendants<Ovml.OleObject>().Any()) {
                        embeddedObjectsBySection[currentSection].Add(new WordEmbeddedObject(_document, run));
                    }
                }

                if (IsSectionBoundaryParagraph(paragraph) && currentSection < embeddedObjectsBySection.Count - 1) {
                    currentSection++;
                }
            }

            return embeddedObjectsBySection[targetSection];
        }

        /// <summary>
        /// Gets list of word elements in given section
        /// </summary>
        /// <returns></returns>
        private List<WordElement> GetWordElements() {
            int targetSection = GetSectionOrdinal();
            var elementsBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<WordElement>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<WordElement>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is Paragraph paragraph) {
                    if (!IsPureSectionBreakParagraph(paragraph)) {
                        elementsBySection[currentSection].AddRange(ConvertParagraphToWordParagraphs(_document, paragraph));
                    }

                    if (IsSectionBoundaryParagraph(paragraph) && currentSection < elementsBySection.Count - 1) {
                        currentSection++;
                    }
                } else if (element is AltChunk altChunk) {
                    elementsBySection[currentSection].Add(new WordEmbeddedDocument(_document, altChunk));
                } else if (element is SdtBlock sdtBlock) {
                    elementsBySection[currentSection].Add(ConvertStdBlockToWordElements(_document, sdtBlock));
                } else if (element is Table table) {
                    elementsBySection[currentSection].Add(new WordTable(_document, table));
                }
            }

            return elementsBySection[targetSection];
        }

        /// <summary>
        /// Gets list of word elements by type in given section
        /// </summary>
        /// <returns></returns>
        private List<WordElement> GetWordElementsByType() {
            var listElements = GetWordElements();
            var additionalElements = new List<WordElement>();

            foreach (var element in listElements) {
                if (element is WordParagraph wordParagraph) {
                    if (wordParagraph.IsBookmark) {
                        additionalElements.Add(wordParagraph.Bookmark!);
                    } else if (wordParagraph.IsBreak) {
                        additionalElements.Add(wordParagraph.Break!);
                    } else if (wordParagraph.IsChart) {
                        additionalElements.Add(wordParagraph.Chart!);
                    } else if (wordParagraph.IsEndNote) {
                        additionalElements.Add(wordParagraph.EndNote!);
                    } else if (wordParagraph.IsEquation) {
                        additionalElements.Add(wordParagraph.Equation!);
                    } else if (wordParagraph.IsField) {
                        additionalElements.Add(wordParagraph.Field!);
                    } else if (wordParagraph.IsFootNote) {
                        additionalElements.Add(wordParagraph.FootNote!);
                    } else if (wordParagraph.IsImage) {
                        additionalElements.Add(wordParagraph.Image!);
                    } else if (wordParagraph.IsListItem) {
                        additionalElements.Add(wordParagraph);
                    } else if (wordParagraph.IsPageBreak) {
                        additionalElements.Add(wordParagraph.PageBreak!);
                    } else if (wordParagraph.IsStructuredDocumentTag) {
                        additionalElements.Add(wordParagraph.StructuredDocumentTag!);
                    } else if (wordParagraph.IsTab) {
                        additionalElements.Add(wordParagraph.Tab!);
                    } else if (wordParagraph.IsTextBox) {
                        additionalElements.Add(wordParagraph.TextBox!);
                    } else if (wordParagraph.IsHyperLink) {
                        additionalElements.Add(wordParagraph.Hyperlink!);
                    } else {
                        additionalElements.Add(wordParagraph);
                    }
                } else {
                    additionalElements.Add(element);
                }
            }
            return additionalElements;
        }

        /// <summary>
        /// Gets list of watermarks in given section
        /// </summary>
        /// <returns></returns>
        private List<SdtBlock> GetSdtBlockList() {
            int targetSection = GetSectionOrdinal();
            var sdtBlocksBySection = Enumerable.Range(0, GetSectionCount())
                .Select(_ => new List<SdtBlock>())
                .ToList();

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return new List<SdtBlock>();
            }

            int currentSection = 0;
            foreach (var element in body.ChildElements) {
                if (element is Paragraph paragraph) {
                    if (IsSectionBoundaryParagraph(paragraph) && currentSection < sdtBlocksBySection.Count - 1) {
                        currentSection++;
                    }
                } else if (element is SdtBlock sdtBlock) {
                    sdtBlocksBySection[currentSection].Add(sdtBlock);
                }
            }

            return sdtBlocksBySection[targetSection];
        }

        /// <summary>
        /// This method moves headers and footers and title page to section before it.
        /// It also copies all other parts of sections (PageSize,PageMargin and others) to section before it.
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
                    } else if (element is FootnoteProperties footnoteProps) {
                        var cloned = (FootnoteProperties)footnoteProps.CloneNode(true);
                        cloned.RemoveAllChildren<NumberingRestart>();
                        newSectionProperties.Append(cloned);
                        footnoteProps.RemoveAllChildren<NumberingRestart>();
                    } else if (element is EndnoteProperties endnoteProps) {
                        var cloned = (EndnoteProperties)endnoteProps.CloneNode(true);
                        cloned.RemoveAllChildren<NumberingRestart>();
                        newSectionProperties.Append(cloned);
                        endnoteProps.RemoveAllChildren<NumberingRestart>();
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
