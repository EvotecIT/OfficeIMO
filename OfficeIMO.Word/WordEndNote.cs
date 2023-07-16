using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    public partial class WordEndNote {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly Run _run;

        public WordEndNote(WordDocument document, Paragraph paragraph, Run run) {
            this._document = document;
            this._paragraph = paragraph;
            this._run = run;
        }

        /// <summary>
        /// List of Paragraphs for given EndNote
        /// As there can be multiple paragraphs with different formatting it's required to provide a list
        /// Zero based object should be skipped, as it's EndNoteReference
        /// However for sake of completion and potential ability to modify it we expose it as well
        /// </summary>
        public List<WordParagraph> Paragraphs {
            get {
                if (_paragraph != null && _run != null) {
                    long referenceId = 0;
                    var endNoteReference = _run.ChildElements.OfType<EndnoteReference>().FirstOrDefault();
                    if (endNoteReference != null) {
                        referenceId = endNoteReference.Id;
                    }

                    if (referenceId != 0) {
                        var endNotesPart = _document._wordprocessingDocument.MainDocumentPart.EndnotesPart;
                        var endNotes = endNotesPart.Endnotes.ChildElements.OfType<Endnote>().ToList();
                        foreach (var endNote in endNotes) {
                            if (endNote != null) {
                                if (endNote.Id == referenceId.ToString()) {
                                    return WordSection.ConvertParagraphsToWordParagraphs(_document, endNote.OfType<Paragraph>());
                                }
                            }
                        }
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Parent Paragraph is Paragraph/Run that has EndNote attached to it.
        /// This provides ability to find proper Run that has EndNote
        /// </summary>

        public WordParagraph ParentParagraph {
            get {
                var previousRun = _run.PreviousSibling<Run>();
                if (previousRun != null) {
                    return new WordParagraph(_document, _paragraph, previousRun);
                }
                return null;
            }
        }

        public long? ReferenceId {
            get {
                if (_paragraph != null && _run != null) {
                    var endNoteReference = _run.ChildElements.OfType<EndnoteReference>().FirstOrDefault();
                    if (endNoteReference != null) {
                        return endNoteReference.Id;
                    }
                }
                return null;
            }
        }

        public void Remove() {
            long referenceId = 0;
            var endNoteReference = _run.ChildElements.OfType<EndnoteReference>().FirstOrDefault();
            if (endNoteReference != null) {
                referenceId = endNoteReference.Id;
            }
            var endNotesPart = _document._wordprocessingDocument.MainDocumentPart.EndnotesPart;
            var footNotes = endNotesPart.Endnotes.ChildElements.OfType<Endnote>().ToList();
            foreach (var footNote in footNotes) {
                if (footNote != null) {
                    if (footNote.Id == referenceId.ToString()) {
                        footNote.Remove();
                    }
                }
            }
            this._run.Remove();
        }


        internal static WordParagraph AddEndNote(WordDocument document, WordParagraph wordParagraph, WordParagraph footerWordParagraph) {

            var endNoteReferenceId = GetNextEndNoteReferenceId(document);

            var newWordParagraph = new WordParagraph(document, wordParagraph._paragraph, true);

            RunStyle runStyle = new RunStyle() { Val = "EndnoteReference" };
            RunProperties runProperties = new RunProperties {
                RunStyle = runStyle
            };
            EndnoteReference endNoteReference = new EndnoteReference() { Id = endNoteReferenceId };
            newWordParagraph._run.Append(runProperties);
            newWordParagraph._run.Append(endNoteReference);

            var endNote = GenerateEndNote(endNoteReferenceId, footerWordParagraph);

            var endNotesPart = document._wordprocessingDocument.MainDocumentPart.EndnotesPart;
            if (endNotesPart == null) {
                endNotesPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<EndnotesPart>();
                WordDocument.GenerateEndNotesPart1Content(endNotesPart);
            }
            endNotesPart.Endnotes.Append(endNote);

            return newWordParagraph;
        }

        internal static long GetNextEndNoteReferenceId(WordDocument document) {
            long highestId = 0;
            var endnotesPart = document._wordprocessingDocument.MainDocumentPart.EndnotesPart;

            // Null check for Endnotes property
            if (endnotesPart?.Endnotes != null) {
                var endNote = endnotesPart.Endnotes.Descendants<Endnote>();

                // Null check for endNote variable
                if (endNote != null && endNote.Any()) {
                    highestId = endNote.Max(en => {
                        if (en.Id != null) {
                            return en.Id.Value;
                        }
                        return 0;
                    });
                } else {
                    highestId = 1;
                }

            }
            return (highestId <= 0) ? 1 : highestId + 1;
        }

        internal static Endnote GenerateEndNote(long endnoteReferenceId, WordParagraph wordParagraph) {
            Endnote endNote = new Endnote() { Id = endnoteReferenceId };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "EndnoteText" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "EndnoteReference" };

            runProperties1.Append(runStyle1);
            EndnoteReferenceMark endnoteReferenceMark = new EndnoteReferenceMark();

            run1.Append(runProperties1);
            run1.Append(endnoteReferenceMark);

            wordParagraph._paragraph.ParagraphProperties = paragraphProperties1;

            var run = wordParagraph._paragraph.GetFirstChild<Run>();
            run.InsertBeforeSelf(run1);

            endNote.Append(wordParagraph._paragraph);

            return endNote;
        }

    }
}
