using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Manages footnotes within a Word document.
    /// </summary>
    public partial class WordFootNote : WordElement {
        private WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly Run _run;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordFootNote"/> class.
        /// </summary>
        public WordFootNote(WordDocument document, Paragraph paragraph, Run run) {
            this._document = document;
            this._paragraph = paragraph;
            this._run = run;
        }

        /// <summary>
        /// List of Paragraphs for given FootNote
        /// As there can be multiple paragraphs with different formatting it's required to provide a list
        /// Zero based object should be skipped, as it's FootnoteReference
        /// However for sake of completion and potential ability to modify it we expose it as well
        /// </summary>
        public List<WordParagraph>? Paragraphs {
            get {
                var footNoteReference = _run.ChildElements.OfType<FootnoteReference>().FirstOrDefault();
                var referenceId = footNoteReference?.Id?.Value ?? 0;

                if (referenceId != 0) {
                    var footnotesPart = _document._wordprocessingDocument.MainDocumentPart?.FootnotesPart;
                    var footNotes = footnotesPart?.Footnotes?.ChildElements.OfType<Footnote>().ToList();
                    if (footNotes != null) {
                        foreach (var footNote in footNotes) {
                            if (footNote?.Id?.Value == referenceId) {
                                return WordSection.ConvertParagraphsToWordParagraphs(_document, footNote.OfType<Paragraph>());
                            }
                        }
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Parent Paragraph is Paragraph/Run that has FootNote attached to it.
        /// This provides ability to find proper Run that has FootNote
        /// </summary>
        public WordParagraph? ParentParagraph {
            get {
                var previousRun = _run.PreviousSibling<Run>();
                if (previousRun != null) {
                    return new WordParagraph(_document, _paragraph, previousRun);
                }
                return null;
            }
        }

        /// <summary>
        /// ReferenceID of FootNote
        /// </summary>
        public long? ReferenceId {
            get {
                var footNoteReference = _run.ChildElements.OfType<FootnoteReference>().FirstOrDefault();
                return footNoteReference?.Id?.Value;
            }
        }

        /// <summary>
        /// Remove FootNote from document
        /// </summary>
        public void Remove() {
            var footNoteReference = _run.ChildElements.OfType<FootnoteReference>().FirstOrDefault();
            var referenceId = footNoteReference?.Id?.Value ?? 0;

            var footnotesPart = _document._wordprocessingDocument.MainDocumentPart?.FootnotesPart;
            var footNotes = footnotesPart?.Footnotes?.ChildElements.OfType<Footnote>().ToList();
            if (footNotes != null) {
                foreach (var footNote in footNotes) {
                    if (footNote?.Id?.Value == referenceId) {
                        footNote.Remove();
                    }
                }
            }

            _run.Remove();
        }

        internal static WordParagraph AddFootNote(WordDocument document, WordParagraph wordParagraph, WordParagraph footerWordParagraph) {
            var footerReferenceId = GetNextFootNotesReferenceId(document);

            var newWordParagraph = new WordParagraph(document, wordParagraph._paragraph, true);

            RunStyle runStyle = new RunStyle() { Val = "FootnoteReference" };
            RunProperties runProperties = new RunProperties {
                RunStyle = runStyle
            };
            FootnoteReference footnoteReference = new FootnoteReference() { Id = footerReferenceId };
            newWordParagraph._run.Append(runProperties);
            newWordParagraph._run.Append(footnoteReference);

            var footNote = GenerateFootNote(footerReferenceId, footerWordParagraph);

            var footNotesPart = document._wordprocessingDocument.MainDocumentPart?.FootnotesPart;
            if (footNotesPart == null) {
                footNotesPart = document._wordprocessingDocument.MainDocumentPart!.AddNewPart<FootnotesPart>();
                WordDocument.GenerateFootNotesPart1Content(footNotesPart);
            }
            footNotesPart.Footnotes!.Append(footNote);

            return newWordParagraph;
        }

        internal static long GetNextFootNotesReferenceId(WordDocument document) {
            long highestId = 0;
            var footnotesPart = document._wordprocessingDocument.MainDocumentPart?.FootnotesPart;

            if (footnotesPart?.Footnotes != null) {
                var footNote = footnotesPart.Footnotes.Descendants<Footnote>();

                if (footNote.Any()) {
                    highestId = footNote.Max(en => en.Id?.Value ?? 0);
                } else {
                    highestId = 1;
                }
            }
            return (highestId <= 0) ? 1 : highestId + 1;
        }

        internal static Footnote GenerateFootNote(long footerReferenceId, WordParagraph wordParagraph) {
            Footnote footnote1 = new Footnote() { Id = footerReferenceId };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "FootnoteText" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "FootnoteReference" };

            runProperties1.Append(runStyle1);
            FootnoteReferenceMark footnoteReferenceMark1 = new FootnoteReferenceMark();

            run1.Append(runProperties1);
            run1.Append(footnoteReferenceMark1);

            wordParagraph._paragraph.ParagraphProperties = paragraphProperties1;

            var run = wordParagraph._paragraph.GetFirstChild<Run>();
            run?.InsertBeforeSelf(run1);

            footnote1.Append(wordParagraph._paragraph);
            return footnote1;
        }
    }
}
