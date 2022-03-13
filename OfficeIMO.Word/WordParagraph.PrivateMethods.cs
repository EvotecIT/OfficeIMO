using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        private Run VerifyRun() {
            if (this._run == null) {
                this._run = new Run();
                this._paragraph.Append(_run);
            }
            return this._run;
        }

        private RunProperties VerifyRunProperties() {
            VerifyRun();
            if (this._run != null) {
                if (this._runProperties == null) {
                    var text = _run.ChildElements.OfType<Text>().FirstOrDefault();
                    if (text != null) {
                        text.InsertBeforeSelf(new RunProperties());
                    } else {
                        this._run.Append(new RunProperties());
                    }
                }
            }
            return this._runProperties;
        }

        private Text VerifyText() {
            if (_text == null) {
                var run = VerifyRun();

                var text = new Text { Space = SpaceProcessingModeValues.Preserve }; // this ensures spaces are preserved between runs 
                this._run.Append(text);
            }

            return this._text;
        }
        private void LoadListToDocument(WordDocument document, WordParagraph wordParagraph) {
            if (wordParagraph.IsListItem) {
                int? listId = wordParagraph._listNumberId;
                if (listId != null) {
                    if (!_document._listNumbersUsed.Contains(listId.Value)) {
                        WordList list = new WordList(wordParagraph._document, document._currentSection, listId.Value);
                        list.ListItems.Add(wordParagraph);
                        _document._listNumbersUsed.Add(listId.Value);
                        _document._currentSection.Lists.Add(list);
                    } else {
                        foreach (WordList list in _document.Lists) {
                            if (list._numberId == listId.Value) {
                                list.ListItems.Add(wordParagraph);
                            }
                        }
                    }
                } else {
                    throw new InvalidOperationException("Couldn't load a list, probably some logic error :-)");
                }
            }
        }

        /// <summary>
        /// Combines the identical runs.
        /// </summary>
        /// <param name="body"></param>
        /// 
        /// 
        /// https://stackoverflow.com/questions/31056953/when-using-mergefield-fieldcodes-in-openxml-sdk-in-c-sharp-why-do-field-codes-di
        public static void CombineIdenticalRuns(Body body) {

            List<Run> runsToRemove = new List<Run>();

            foreach (Paragraph para in body.Descendants<Paragraph>()) {
                List<Run> runs = para.Elements<Run>().ToList();
                for (int i = runs.Count - 2; i >= 0; i--) {
                    Text text1 = runs[i].GetFirstChild<Text>();
                    Text text2 = runs[i + 1].GetFirstChild<Text>();
                    if (text1 != null && text2 != null) {
                        string rPr1 = "";
                        string rPr2 = "";
                        if (runs[i].RunProperties != null) rPr1 = runs[i].RunProperties.OuterXml;
                        if (runs[i + 1].RunProperties != null) rPr2 = runs[i + 1].RunProperties.OuterXml;
                        if (rPr1 == rPr2) {
                            text1.Text += text2.Text;
                            runsToRemove.Add(runs[i + 1]);
                        }
                    }
                }
            }
            foreach (Run run in runsToRemove) {
                run.Remove();
            }
        }
    }
}
