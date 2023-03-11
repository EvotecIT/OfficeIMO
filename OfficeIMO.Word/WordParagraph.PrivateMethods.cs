using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;

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
        internal Run VerifyRun() {
            if (this._run == null) {
                this._run = new Run();
                this._paragraph.Append(_run);
            }

            return this._run;
        }

        internal Run VerifyRun(Paragraph paragraph, Run run) {
            if (run == null) {
                run = new Run();
                paragraph.Append(run);
            }

            return run;
        }

        internal Run VerifyRun(Hyperlink hyperlink, Run run) {
            if (run == null) {
                run = new Run();
                hyperlink.Append(run);
            }

            return run;
        }

        private RunProperties VerifyRunProperties(Hyperlink hyperlink, Run run, RunProperties runProperties) {
            VerifyRun(hyperlink, run);
            if (run != null) {
                if (runProperties == null) {
                    var text = run.ChildElements.OfType<Text>().FirstOrDefault();
                    if (text != null) {
                        text.InsertBeforeSelf(new RunProperties());
                    } else {
                        run.Append(new RunProperties());
                    }
                }
            }

            return runProperties;
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

        /// <summary>
        /// Returns a Text field. If it doesn't exits creates it.
        /// </summary>
        /// <returns></returns>
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

        private void ParseTextForOpenXml(Run run, string text) {
            //string[] newLineArray = { Environment.NewLine };
            string[] newLineArray = { Environment.NewLine, "\n", "\r\n", "\n\r" };
            string[] textArray = text.Split(newLineArray, StringSplitOptions.None);

            bool first = true;

            foreach (string line in textArray) {
                if (!first) {
                    run.Append(new Break());
                }

                first = false;

                Text txt = new Text {
                    Text = line
                };

                run.Append(txt);
            }
        }

        private List<string> ConvertStringToList(string text) {
            string[] splitStrings = { Environment.NewLine, "\r\n", "\n" };
            string[] textSplit = text.Split(splitStrings, StringSplitOptions.RemoveEmptyEntries);
            var list = new List<string>();
            for (int i = 0; i < textSplit.Length; i++) {
                // check if there's new line at the beginning of the text
                // if there is add empty string to the list
                if (i == 0 && text.StartsWith(Environment.NewLine)) {
                    list.Add("");
                } else if (i == 0 && text.StartsWith("\r\n")) {
                    list.Add("");
                } else if (i == 0 && text.StartsWith("\n")) {
                    list.Add("");
                }
                // add splitted text to the list
                list.Add(textSplit[i]);

                if (i < textSplit.Length - 1) {
                    // for every element in the list except the last element add empty string to the list
                    list.Add("");
                } else {
                    // check if there's new line at the end of the text
                    // if there is add an empty string to the list
                    if (text.EndsWith(Environment.NewLine)) {
                        list.Add("");
                    } else if (text.EndsWith("\r\n")) {
                        list.Add("");
                    } else if (text.EndsWith("\n")) {
                        list.Add("");
                    }
                }
            }
            return list;
        }

        private WordParagraph ConvertToTextWithBreaks(string text) {
            string[] splitStrings = { Environment.NewLine, "\r\n", "\n" };

            WordParagraph wordParagraph = null;

            // check if there's a new line in the text
            if (splitStrings.Any(text.Contains)) {
                // if there is new line in the text, split the text and add a new paragraph for each line
                // for any new line, add a break
                var listOfText = ConvertStringToList(text);
                foreach (string line in listOfText) {
                    if (line == "") {
                        wordParagraph = AddBreak();
                    } else {
                        wordParagraph = new WordParagraph(this._document, this._paragraph, new Run());
                        wordParagraph.Text = line;
                        this._paragraph.Append(wordParagraph._run);
                    }
                }
            } else {
                wordParagraph = new WordParagraph(this._document, this._paragraph, new Run());
                wordParagraph.Text = text;
                this._paragraph.Append(wordParagraph._run);
            }

            return wordParagraph;
        }
    }
}