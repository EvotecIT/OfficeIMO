using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using SemanticComparison;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_CreatingWordWithFootNotes() {
            using (WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "FootNotesEndNotes.docx"))) {
                Assert.True(document.Paragraphs.Count == 0);
                Assert.True(document.Sections.Count == 1);
                Assert.True(document.Fields.Count == 0);
                Assert.True(document.HyperLinks.Count == 0);
                Assert.True(document.ParagraphsHyperLinks.Count == 0);
                Assert.True(document.Bookmarks.Count == 0);

                var paragraph = document.AddParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                document.AddParagraph("This is my text").AddFootNote("This is a footnote to my text").AddText(" continuing").AddFootNote("2nd footnote!");

                Assert.True(document.FootNotes.Count == 2);
                Assert.True(document.Sections[0].FootNotes.Count == 2);
                Assert.True(document.Sections[0].EndNotes.Count == 0);


                var lastFootNoteParagraph = document.AddParagraph("Another paragraph").AddFootNote("more footnotes!")
                    .AddText(" more within paragraph").AddFootNote("4th footnote!");

                Assert.True(document.FootNotes.Count == 4);
                Assert.True(document.Sections[0].FootNotes.Count == 4);
                Assert.True(document.Sections[0].EndNotes.Count == 0);

                Assert.True(lastFootNoteParagraph.IsFootNote == true);

                var footNoteParagraphs = lastFootNoteParagraph.FootNote.Paragraphs;

                Assert.True(footNoteParagraphs.Count == 2);

                Assert.True(lastFootNoteParagraph.FootNote.ParentParagraph.Text == " more within paragraph");
                Assert.True(footNoteParagraphs[1].Text == "4th footnote!");


                Assert.True(document.FootNotes[3].ParentParagraph.Text == " more within paragraph");
                Assert.True(document.FootNotes[3].Paragraphs[1].Text == "4th footnote!");

                Assert.True(document.FootNotes[3].Paragraphs[1].Bold == false);

                // lets make bold that footnote
                footNoteParagraphs[1].Bold = true;

                Assert.True(footNoteParagraphs[1].Bold == true);
                Assert.True(document.FootNotes[3].Paragraphs[1].Bold == true);

                document.AddParagraph("Testing endnote - 1").AddEndNote("Test end note 1");

                Assert.True(document.EndNotes.Count == 1);
                Assert.True(document.FootNotes.Count == 4);
                Assert.True(document.Sections[0].FootNotes.Count == 4);
                Assert.True(document.Sections[0].EndNotes.Count == 1);

                document.AddParagraph("Test 1");

                document.AddSection();

                document.AddParagraph("Testing endnote - 2").AddEndNote("Test end note 2");

                Assert.True(document.EndNotes.Count == 2);
                Assert.True(document.FootNotes.Count == 4);
                Assert.True(document.Sections[0].FootNotes.Count == 4);
                Assert.True(document.Sections[0].EndNotes.Count == 1);
                Assert.True(document.Sections[1].EndNotes.Count == 1);


                document.AddParagraph("Another paragraph 1").AddFootNote("more footnotes 2!");

                Assert.True(document.EndNotes.Count == 2);
                Assert.True(document.FootNotes.Count == 5);
                Assert.True(document.Sections[0].FootNotes.Count == 4);
                Assert.True(document.Sections[1].FootNotes.Count == 1);
                Assert.True(document.Sections[0].EndNotes.Count == 1);
                Assert.True(document.Sections[1].EndNotes.Count == 1);

                document.FootNotes[1].Remove();

                Assert.True(document.EndNotes.Count == 2);
                Assert.True(document.FootNotes.Count == 4);
                Assert.True(document.Sections[0].FootNotes.Count == 3);
                Assert.True(document.Sections[1].FootNotes.Count == 1);
                Assert.True(document.Sections[0].EndNotes.Count == 1);
                Assert.True(document.Sections[1].EndNotes.Count == 1);

                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "FootNotesEndNotes.docx"))) {
                Assert.True(document.EndNotes.Count == 2);
                Assert.True(document.FootNotes.Count == 4);
                Assert.True(document.Sections[0].FootNotes.Count == 3);
                Assert.True(document.Sections[1].FootNotes.Count == 1);
                Assert.True(document.Sections[0].EndNotes.Count == 1);
                Assert.True(document.Sections[1].EndNotes.Count == 1);

                document.AddParagraph("This is my text").AddFootNote("This is a footnote to my text").AddText(" continuing").AddFootNote("2nd footnote!");

                Assert.True(document.EndNotes.Count == 2);
                Assert.True(document.FootNotes.Count == 6);
                Assert.True(document.Sections[0].FootNotes.Count == 3);
                Assert.True(document.Sections[1].FootNotes.Count == 3);
                Assert.True(document.Sections[0].EndNotes.Count == 1);
                Assert.True(document.Sections[1].EndNotes.Count == 1);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "FootNotesEndNotes.docx"))) {

                Assert.True(document.EndNotes.Count == 2);
                Assert.True(document.FootNotes.Count == 6);
                Assert.True(document.Sections[0].FootNotes.Count == 3);
                Assert.True(document.Sections[1].FootNotes.Count == 3);
                Assert.True(document.Sections[0].EndNotes.Count == 1);
                Assert.True(document.Sections[1].EndNotes.Count == 1);

                document.Save();
            }
        }

        [Fact]
        public void Test_LoadingWordWithFootNotesAndEndNotes() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "DocumentWithFootNotes.docx"))) {

                Assert.True(document.EndNotes.Count == 2);
                Assert.True(document.FootNotes.Count == 3);
                Assert.True(document.Sections[0].FootNotes.Count == 3);
                Assert.True(document.Sections[0].EndNotes.Count == 2);

                document.AddParagraph("This is my text").AddFootNote("This is a footnote to my text").AddText(" continuing").AddFootNote("2nd footnote!");

                Assert.True(document.EndNotes.Count == 2);
                Assert.True(document.FootNotes.Count == 5);
                Assert.True(document.Sections[0].FootNotes.Count == 5);
                Assert.True(document.Sections[0].EndNotes.Count == 2);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "DocumentWithFootNotes.docx"))) {
                Assert.True(document.EndNotes.Count == 2);
                Assert.True(document.FootNotes.Count == 5);
                Assert.True(document.Sections[0].FootNotes.Count == 5);
                Assert.True(document.Sections[0].EndNotes.Count == 2);

            }
        }


        [Fact]
        public void Test_LoadingEmptyWordAndAddingFootNotesEndNotes() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "EmptyDocument.docx"))) {
                Assert.True(document.EndNotes.Count == 0);
                Assert.True(document.FootNotes.Count == 0);
                Assert.True(document.Sections[0].FootNotes.Count == 0);
                Assert.True(document.Sections[0].EndNotes.Count == 0);

                document.AddParagraph("This is my text").AddFootNote("This is a footnote to my text").AddText(" continuing").AddFootNote("2nd footnote!");

                Assert.True(document.EndNotes.Count == 0);
                Assert.True(document.FootNotes.Count == 2);
                Assert.True(document.Sections[0].FootNotes.Count == 2);
                Assert.True(document.Sections[0].EndNotes.Count == 0);

                document.AddParagraph("Testing endnote - 2").AddEndNote("Test end note 2");

                Assert.True(document.EndNotes.Count == 1);
                Assert.True(document.FootNotes.Count == 2);
                Assert.True(document.Sections[0].FootNotes.Count == 2);
                Assert.True(document.Sections[0].EndNotes.Count == 1);

                var filePath = Path.Combine(_directoryWithFiles, "DocumentWithFootNotes1.docx");

                document.Save(filePath);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "DocumentWithFootNotes1.docx"))) {
                Assert.True(document.EndNotes.Count == 1);
                Assert.True(document.FootNotes.Count == 2);
                Assert.True(document.Sections[0].FootNotes.Count == 2);
                Assert.True(document.Sections[0].EndNotes.Count == 1);

                document.Save(false);
            }
        }
    }
}
