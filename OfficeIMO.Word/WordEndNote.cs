using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    public partial class WordEndNote {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly List<Run> _runs = new List<Run>();

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
