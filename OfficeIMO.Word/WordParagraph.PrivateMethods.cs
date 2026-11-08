using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // Contains private helper methods for WordParagraph.
    public partial class WordParagraph {
        /// <summary>
        /// Checks which document story owns the paragraph.
        /// This is required for image processing because each story owns its own image relationships
        /// as the location of the image matters to be able to display it properly.
        /// </summary>
        /// <returns></returns>
        internal OpenXmlElement Location() {
            var parent = this._paragraph.Parent;

            while (parent != null) {
                if (parent is Header
                    or Footer
                    or Document
                    or Footnotes
                    or Endnotes
                    or Comments) {
                    return parent;
                }

                parent = parent.Parent;
            }

            throw new InvalidOperationException("Unable to determine the paragraph's document story.");
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

        private RunProperties VerifyRunProperties(Hyperlink hyperlink, Run run, RunProperties? runProperties) {
            run = VerifyRun(hyperlink, run);
            if (run != null) {
                runProperties = run.GetFirstChild<RunProperties>();
                if (runProperties == null) {
                    runProperties = run.PrependChild(new RunProperties());
                }
            }

            return runProperties!;
        }

        /// <summary>
        /// Returns the run properties, creating the run and its first child when needed.
        /// </summary>
        /// <returns></returns>
        internal RunProperties VerifyRunProperties() {
            if (_stdRun != null) {
                SdtContentRun content = _stdRun.SdtContentRun ??= new SdtContentRun();
                Run run = content.Descendants<Run>().FirstOrDefault() ?? content.AppendChild(new Run());
                return run.GetFirstChild<RunProperties>() ?? run.PrependChild(new RunProperties());
            }
            VerifyRun();
            if (_run == null) {
                throw new InvalidOperationException("Run is not initialized.");
            }

            var runProperties = _run.GetFirstChild<RunProperties>();
            if (runProperties == null) {
                runProperties = _run.PrependChild(new RunProperties());
            }

            return runProperties;
        }

        /// <summary>
        /// Returns a Text field. If it doesn't exits creates it.
        /// </summary>
        /// <returns></returns>
        private Text VerifyText() {
            if (_text == null) {
                var run = VerifyRun();
                var text = new Text { Space = SpaceProcessingModeValues.Preserve };
                run.Append(text);
                return text;
            }
            return this._text!;
        }

        //private void LoadListToDocument(WordDocument document, WordParagraph wordParagraph) {
        //    if (wordParagraph.IsListItem) {
        //        int? listId = wordParagraph._listNumberId;
        //        if (listId != null) {
        //            if (!_document._listNumbersUsed.Contains(listId.Value)) {
        //                WordList list = new WordList(wordParagraph._document, document._currentSection, listId.Value);
        //                list.ListItems.Add(wordParagraph);
        //                _document._listNumbersUsed.Add(listId.Value);
        //                _document._currentSection.Lists.Add(list);
        //            } else {
        //                foreach (WordList list in _document.Lists) {
        //                    if (list._numberId == listId.Value) {
        //                        list.ListItems.Add(wordParagraph);
        //                    }
        //                }
        //            }
        //        } else {
        //            throw new InvalidOperationException("Couldn't load a list, probably some logic error :-)");
        //        }
        //    }
        //}
        private WordParagraph ConvertToTextWithBreaks(string text) {
            if (text.IndexOf('\n') < 0) {
                var textParagraph = new WordParagraph(this._document, this._paragraph, new Run(), Parent);
                textParagraph.Text = text;
                this._paragraph.Append(textParagraph._run!);
                return textParagraph;
            }

            string[] lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            WordParagraph wordParagraph = this;
            for (int i = 0; i < lines.Length; i++) {
                if (i > 0) {
                    wordParagraph = AddBreak();
                }
                if (lines[i].Length > 0) {
                    wordParagraph = new WordParagraph(this._document, this._paragraph, new Run(), Parent);
                    wordParagraph.Text = lines[i];
                    this._paragraph.Append(wordParagraph._run!);
                }
            }
            return wordParagraph;
        }
    }
}
