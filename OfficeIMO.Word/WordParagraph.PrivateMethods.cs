using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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
    }
}
