using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Checks where the paragraph is located. If it is located in the header, footer or main document.
        /// This is required for the image processing to work properly for header and footers
        /// as the location of the image matters to be able to display it properly.
        /// </summary>
        /// <returns></returns>
        internal OpenXmlElement Location() {
            // i'm assuming the depth shouldn't be more than 10 to get a parent of paragraph
            int count = 0;
            var parent = this._paragraph.Parent;

            do {
                if (parent != null) {
                    if (parent.GetType() == typeof(Header)) {
                        return parent;
                    } else if (parent.GetType() == typeof(Footer)) {
                        return parent;
                    } else if (parent.GetType() == typeof(Document)) {
                        return parent;
                    }

                    parent = parent.Parent;
                }
                count++;
            } while (count < 10 || parent != null);

            return null;
        }

        /// <summary>
        /// Check if run exists, if not create it and append to paragraph
        /// </summary>
        /// <returns></returns>
        private Run VerifyRun() {
            if (this._run == null) {
                this._run = new Run();
                this._paragraph.Append(_run);
            }
            return this._run;
        }

        /// <summary>
        /// Check if runProperties exists in run, if not create run, create run properties and and append to run
        /// </summary>
        /// <returns></returns>
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
