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
        private readonly FootnoteReference? _reference;

        /// <summary>
        /// Initializes a note handle for an exact reference child, or the run's first reference
        /// when a paragraph's singular note property creates the handle.
        /// </summary>
        internal WordFootNote(WordDocument document, Paragraph paragraph, Run run, FootnoteReference? reference = null) {
            this._document = document;
            this._paragraph = paragraph;
            this._run = run;
            this._reference = reference ?? run.Elements<FootnoteReference>().FirstOrDefault();
        }

        /// <summary>
        /// Gets the paragraphs of the referenced ordinary footnote, or null when its body cannot be resolved.
        /// A note can contain multiple paragraphs with different formatting.
        /// </summary>
        public List<WordParagraph>? Paragraphs {
            get {
                Footnote? footnote = WordNoteResolver.FindFootnote(
                    _document._wordprocessingDocument.MainDocumentPart, _reference?.Id?.Value);
                return footnote == null ? null : WordSection.ConvertParagraphsToWordParagraphs(_document, footnote.Elements<Paragraph>());
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
                return _reference?.Id?.Value;
            }
        }

        /// <summary>
        /// Removes this footnote and its reference, preserving other content in the reference run.
        /// </summary>
        public void Remove() {
            WordNoteResolver.FindFootnote(
                _document._wordprocessingDocument.MainDocumentPart, _reference?.Id?.Value)?.Remove();

            if (_reference?.Parent == _run) {
                _reference.Remove();
                if (_run.Parent != null && _run.ChildElements.All(child => child is RunProperties)) {
                    _run.Remove();
                }
            }
        }

        internal static WordParagraph AddFootNote(WordDocument document, WordParagraph wordParagraph, WordParagraph footerWordParagraph) {
            if (wordParagraph._paragraph.Ancestors().Any(ancestor => ancestor is Header || ancestor is Footer)) {
                throw new InvalidOperationException("Footnotes cannot be added to header or footer paragraphs.");
            }

            var footerReferenceId = GetNextFootNotesReferenceId(document);

            var newWordParagraph = new WordParagraph(document, wordParagraph._paragraph, true);

            RunStyle runStyle = new RunStyle() { Val = "FootnoteReference" };
            RunProperties runProperties = new RunProperties {
                RunStyle = runStyle
            };
            FootnoteReference footnoteReference = new FootnoteReference() { Id = footerReferenceId };
            if (newWordParagraph._run == null) {
                newWordParagraph._run = new Run();
                newWordParagraph._paragraph.Append(newWordParagraph._run);
            }
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

            ParagraphProperties paragraphProperties1 = wordParagraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraphProperties1.ParagraphStyleId ??= new ParagraphStyleId { Val = "FootnoteText" };

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "FootnoteReference" };

            runProperties1.Append(runStyle1);
            FootnoteReferenceMark footnoteReferenceMark1 = new FootnoteReferenceMark();

            run1.Append(runProperties1);
            run1.Append(footnoteReferenceMark1);

            wordParagraph._paragraph.InsertAfter(run1, paragraphProperties1);

            footnote1.Append(wordParagraph._paragraph);
            wordParagraph.RefreshParent();
            return footnote1;
        }
    }
}
