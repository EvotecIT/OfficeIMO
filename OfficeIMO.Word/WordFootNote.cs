using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    public partial class WordFootNote {
        private WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly Run _run;

        public WordFootNote(WordDocument document, Paragraph paragraph, Run run) {
            this._document = document;
            this._paragraph = paragraph;
            this._run = run;
        }

        //public void Remove(bool includingParagraph = false) {
        //    if (includingParagraph) {
        //        this._paragraph.Remove();
        //    } else {
        //        if (_run.ChildElements.Count == 1) {
        //            this._run.Remove();
        //        } else {
        //            this._run.ChildElements.OfType<TabChar>().FirstOrDefault()?.Remove();
        //        }
        //    }
        //}

        //public List<WordParagraph> Paragraphs {
        //    get {
        //        // Get the FootnotesPart
        //        FootnotesPart footnotesPart = _document._wordprocessingDocument.MainDocumentPart.FootnotesPart;

        //        foreach (Footnote footnote in footnotesPart.Footnotes) {
        //            // Find the last Paragraph element in the Footnote element.
        //            Paragraph lastParagraph = footnote.Descendants<Paragraph>().LastOrDefault();
        //            if (lastParagraph != null) {
        //                // Get the Text from that Paragraph element
        //                IEnumerable<Text> paragraphText = lastParagraph.Descendants<Text>();
        //            }
        //        }

        //    };
        //}

        //public Footnote GetFootnoteById(WordprocessingDocument wordprocessingDocument, string id) {
        //    // Get the FootnotesPart
        //    FootnotesPart footnotesPart = wordprocessingDocument.MainDocumentPart.FootnotesPart;

        //    // Search for the footnote by ID
        //    foreach (Footnote footnote in footnotesPart.Footnotes) {
        //        if (footnote.Id == id) {
        //            return footnote;
        //        }
        //    }

        //    return null;
        //}




        internal static WordParagraph AddFootNote(WordDocument document, WordParagraph wordParagraph, WordParagraph footerWordParagraph) {
            // _document = document;
            // _paragraph = wordParagraph._paragraph;


            var footerReferenceId = GetNextFooterReferenceId(document);

            var newWordParagraph = new WordParagraph(document, wordParagraph._paragraph, true);

            RunStyle runStyle = new RunStyle() { Val = "FootnoteReference" };
            RunProperties runProperties = new RunProperties {
                RunStyle = runStyle
            };
            FootnoteReference footnoteReference = new FootnoteReference() { Id = footerReferenceId };
            newWordParagraph._run.Append(runProperties);
            newWordParagraph._run.Append(footnoteReference);

            var footNote = GenerateFootNote(footerReferenceId, footerWordParagraph);

            var footNotesPart = document._wordprocessingDocument.MainDocumentPart.FootnotesPart;
            footNotesPart.Footnotes.Append(footNote);

            return newWordParagraph;
        }



        internal static long GetNextFooterReferenceId(WordDocument document) {
            var footNotesPart = document._wordprocessingDocument.MainDocumentPart.FootnotesPart;
            var highestId = footNotesPart.Footnotes.Descendants<Footnote>().Max(fn => fn.Id.Value);
            return (highestId <= 0) ? 1 : highestId + 1;
        }


        internal static Footnote GenerateFootNote(long footerReferenceId, WordParagraph wordParagraph) {
            Footnote footnote1 = new Footnote() { Id = footerReferenceId };

            //  Paragraph paragraph1 = new Paragraph() { };

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
            // wordParagraph._run.RunProperties = runProperties1;

            var run = wordParagraph._paragraph.GetFirstChild<Run>();
            run.InsertBeforeSelf(run1);

            //var text = wordParagraph._run.GetFirstChild<Text>();
            //text.InsertBeforeSelf(footnoteReferenceMark1);
            //wordParagraph._run.Append(footnoteReferenceMark1);


            footnote1.Append(wordParagraph._paragraph);

            //Run run2 = new Run();
            //Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            //text1.Text = " This is first ";

            //run2.Append(text1);
            //ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            //Run run3 = new Run();
            //Text text2 = new Text();
            //text2.Text = "footnote";

            //run3.Append(text2);
            //ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            //paragraph1.Append(paragraphProperties1);
            //paragraph1.Append(run1);
            //paragraph1.Append(run2);
            //paragraph1.Append(proofError1);
            //paragraph1.Append(run3);
            //paragraph1.Append(proofError2);

            //footnote1.Append(paragraph1);
            return footnote1;
        }




    }
}
