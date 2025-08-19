using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    /// <summary>
    /// Handles endnotes.
    /// </summary>
    public partial class WordEndNote : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly Run _run;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordEndNote"/> class.
        /// </summary>
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
        public List<WordParagraph>? Paragraphs {
            get {
                long referenceId = _run.ChildElements.OfType<EndnoteReference>().FirstOrDefault()?.Id?.Value ?? 0;

                if (referenceId != 0) {
                    var endNotesPart = _document._wordprocessingDocument.MainDocumentPart?.EndnotesPart;
                    var endNotes = endNotesPart?.Endnotes?.ChildElements.OfType<Endnote>().ToList();
                    if (endNotes != null) {
                        foreach (var endNote in endNotes) {
                            if (endNote != null && endNote.Id == referenceId.ToString()) {
                                return WordSection.ConvertParagraphsToWordParagraphs(_document, endNote.OfType<Paragraph>());
                            }
                        }
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Parent paragraph containing the endnote reference.
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
        /// Gets the endnote reference identifier if available.
        /// </summary>
        public long? ReferenceId {
            get {
                var endNoteReference = _run.ChildElements.OfType<EndnoteReference>().FirstOrDefault();
                return endNoteReference?.Id?.Value;
            }
        }

        /// <summary>
        /// Removes the endnote and its reference from the document.
        /// </summary>
        public void Remove() {
            long referenceId = _run.ChildElements.OfType<EndnoteReference>().FirstOrDefault()?.Id?.Value ?? 0;
            var endNotesPart = _document._wordprocessingDocument.MainDocumentPart?.EndnotesPart;
            var footNotes = endNotesPart?.Endnotes?.ChildElements.OfType<Endnote>().ToList();
            if (footNotes != null) {
                foreach (var footNote in footNotes) {
                    if (footNote != null && footNote.Id == referenceId.ToString()) {
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

            var mainDocumentPart = document._wordprocessingDocument.MainDocumentPart ?? throw new InvalidOperationException("Document missing MainDocumentPart");
            var endNotesPart = mainDocumentPart.EndnotesPart;
            if (endNotesPart == null) {
                endNotesPart = mainDocumentPart.AddNewPart<EndnotesPart>();
                WordDocument.GenerateEndNotesPart1Content(endNotesPart);
            }
            endNotesPart.Endnotes!.Append(endNote);

            return newWordParagraph;
        }

        internal static long GetNextEndNoteReferenceId(WordDocument document) {
            long highestId = 0;
            var endnotesPart = document._wordprocessingDocument.MainDocumentPart?.EndnotesPart;

            if (endnotesPart?.Endnotes != null) {
                var endNote = endnotesPart.Endnotes.Descendants<Endnote>();

                if (endNote.Any()) {
                    highestId = endNote.Max(en => en.Id?.Value ?? 0);
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
            run?.InsertBeforeSelf(run1);

            endNote.Append(wordParagraph._paragraph);

            return endNote;
        }

    }
}
