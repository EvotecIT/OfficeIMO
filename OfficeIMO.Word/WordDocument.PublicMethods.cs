using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        public WordParagraph AddParagraph(WordParagraph wordParagraph = null) {
            if (wordParagraph == null) {
                // we create paragraph (and within that add it to document)
                wordParagraph = new WordParagraph(this);
            } else {
                // since we created paragraph without adding it to document, we now need to add it to document
                //this.Paragraphs.Add(wordParagraph);
            }

            //this._currentSection.Paragraphs.Add(wordParagraph);
            // wordParagraph._section = this._currentSection;
            this._wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(wordParagraph._paragraph);
            return wordParagraph;
        }

        public WordParagraph AddParagraph(string text) {
            return AddParagraph().SetText(text);
        }

        public WordParagraph AddPageBreak() {
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() { Type = BreakValues.Page }),
                _document = this
            };
            newWordParagraph._paragraph = new Paragraph(newWordParagraph._run);

            this._document.Body.Append(newWordParagraph._paragraph);
            return newWordParagraph;
        }

        public WordParagraph AddBreak(BreakValues breakType = BreakValues.Page) {
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() { Type = breakType }),
                _document = this
            };
            newWordParagraph._paragraph = new Paragraph(newWordParagraph._run);

            this._document.Body.Append(newWordParagraph._paragraph);
            this.Paragraphs.Add(newWordParagraph);
            return newWordParagraph;
        }

        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, uri, addStyle, tooltip, history);
        }

        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, anchor, addStyle, tooltip, history);
        }
        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this, this._currentSection);
            wordList.AddList(style);
            return wordList;
        }

        public WordList AddTableOfContentList(WordListStyle style) {
            WordList wordList = new WordList(this, this._currentSection, true);
            wordList.AddList(style);
            return wordList;
        }

        public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this, this._currentSection, rows, columns, tableStyle);
            return wordTable;
        }

        public WordTableOfContent AddTableOfContent(TableOfContentStyle tableOfContentStyle = TableOfContentStyle.Template1) {
            WordTableOfContent wordTableContent = new WordTableOfContent(this, tableOfContentStyle);
            return wordTableContent;
        }

        public WordParagraph AddHorizontalLine(BorderValues lineType = BorderValues.Single, System.Drawing.Color? color = null, uint size = 12, uint space = 1) {
            return this.AddParagraph().AddHorizontalLine(lineType, color, size, space);
        }

        public WordSection AddSection(SectionMarkValues? sectionMark = null) {
            //Paragraph paragraph = new Paragraph() { RsidParagraphAddition = "fff0", RsidRunAdditionDefault = "fff0"};
            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProperties = new ParagraphProperties();

            SectionProperties sectionProperties = new SectionProperties();
            // SectionProperties sectionProperties = new SectionProperties() { RsidR = "fff0"  };

            if (sectionMark != null) {
                SectionType sectionType = new SectionType() { Val = sectionMark };
                sectionProperties.Append(sectionType);
            }

            paragraphProperties.Append(sectionProperties);
            paragraph.Append(paragraphProperties);


            this._document.MainDocumentPart.Document.Body.Append(paragraph);


            WordSection wordSection = new WordSection(this, paragraph);

            return wordSection;
        }

        public List<WordParagraph> FindText(string text) {
            throw new NotImplementedException("It's not implemented :-)");
            List<WordParagraph> list = new List<WordParagraph>();


            return list;
            //Paragraph currentParagraph = null;
            //foreach (var paragraph in this.Paragraphs) {
            //    if (currentParagraph == null) {
            //        currentParagraph = paragraph._paragraph;
            //    } else {
            //        if (currentParagraph != paragraph._paragraph) {

            //        }
            //    }


            //    // paragraph.Text
            //    // logic missing to find text that can start within some paragraph.Text, but can span across multiple paragraphs
            //    // for example searching for text "This Is MyTest" within 4 paragraphs that may be written like
            //    // paragraph.Text = "Thi"
            //    // paragraph.Text = "s Is"
            //    // paragraph.Text = " MyTes"
            //    // paragraph.Text = "t"


            //}
            //Dictionary<int, List<WordParagraph>> dic = new Dictionary<int, List<WordParagraph>>();
            //Paragraph tempParagraph = null;
            //bool sameParagraph = false;
            //int count = 0;

            //dic[count] = new List<WordParagraph>();
            //for (int i = 0; i < this.Paragraphs.Count - 1; i++) {
            //    var paragraph = this.Paragraphs[i];
            //    if (tempParagraph == null || tempParagraph == paragraph._paragraph) {
            //        sameParagraph = true;
            //        tempParagraph = paragraph._paragraph;
            //    } else {

            //        sameParagraph = false;
            //        count++;
            //        dic[count] = new List<WordParagraph>();
            //    }

            //    if (sameParagraph) {
            //        dic[count].Add(paragraph);

            //    }
            //}

            //foreach (int i in dic.Keys) {
            //    //dic[i]
            //}


            //return list;
        }
    }
}
