using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_OpeningWordWithFieldsAndHyperlinks() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "FieldsAndSections.docx"))) {
                Assert.True(document.Paragraphs.Count == 30);
                Assert.True(document.Sections.Count == 2);
                Assert.True(document.Fields.Count == 3);
                Assert.True(document.HyperLinks.Count == 2);

                Assert.True(document.Sections[0].Fields.Count == 3);
                Assert.True(document.Sections[0].HyperLinks.Count == 2);

                Assert.True(document.Sections[1].Fields.Count == 0);
                Assert.True(document.Sections[1].HyperLinks.Count == 0);

                Assert.True(document.ParagraphsHyperLinks[0].Hyperlink.IsEmail == false);
                Assert.True(document.ParagraphsHyperLinks[1].Hyperlink.IsEmail == true);
                Assert.True(document.ParagraphsHyperLinks[1].Hyperlink.EmailAddress == "przemyslaw.klys@test.pl");

                Assert.True(document.Sections[0].ParagraphsFields[0].Field.Text == "Przemysław Kłys");
                Assert.True(document.Sections[0].ParagraphsFields[1].Field.Text == "FieldsAndSections.docx");
                Assert.True(document.Sections[0].ParagraphsFields[2].Field.Text == "1");

                Assert.True(document.Sections[0].ParagraphsFields[0].Field.Field == @" AUTHOR  \* Caps  \* MERGEFORMAT ");
                Assert.True(document.Sections[0].ParagraphsFields[1].Field.Field == @" FILENAME   \* MERGEFORMAT ");
                Assert.True(document.Sections[0].ParagraphsFields[2].Field.Field == @" PAGE  \* Arabic  \* MERGEFORMAT ");
            }
        }
        [Fact]
        public void Test_OpeningWordWithFieldsAndHyperlinksAndEquationsAndBookmarks() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "FieldsAndSectionsAdvanced.docx"))) {
                Assert.True(document.Paragraphs.Count == 52);
                Assert.True(document.Sections.Count == 2);
                Assert.True(document.Fields.Count == 3);
                Assert.True(document.HyperLinks.Count == 2);
                Assert.True(document.Equations.Count == 2);
                Assert.True(document.StructuredDocumentTags.Count == 2);
                Assert.True(document.Bookmarks.Count == 3);
                Assert.True(document.Comments.Count == 0);
                Assert.True(document.Images.Count == 0);
                Assert.True(document.PageBreaks.Count == 1);

                Assert.True(document.ParagraphsFields.Count == 3);
                Assert.True(document.ParagraphsHyperLinks.Count == 2);
                Assert.True(document.ParagraphsEquations.Count == 2);
                Assert.True(document.ParagraphsStructuredDocumentTags.Count == 2);
                Assert.True(document.ParagraphsBookmarks.Count == 3);
                Assert.True(document.Comments.Count == 0);
                Assert.True(document.ParagraphsImages.Count == 0);
                Assert.True(document.ParagraphsPageBreaks.Count == 1);

                Assert.True(document.Sections[0].ParagraphsFields.Count == 3);
                Assert.True(document.Sections[0].ParagraphsHyperLinks.Count == 2);
                Assert.True(document.Sections[0].ParagraphsEquations.Count == 0);
                Assert.True(document.Sections[0].ParagraphsStructuredDocumentTags.Count == 0);
                Assert.True(document.Sections[0].ParagraphsBookmarks.Count == 2);
                //Assert.True(document.Sections[0].Comments.Count == 0);
                Assert.True(document.Sections[0].ParagraphsImages.Count == 0);
                Assert.True(document.Sections[0].ParagraphsPageBreaks.Count == 1);

                Assert.True(document.Sections[1].ParagraphsFields.Count == 0);
                Assert.True(document.Sections[1].ParagraphsHyperLinks.Count == 0);
                Assert.True(document.Sections[1].ParagraphsEquations.Count == 2);
                Assert.True(document.Sections[1].ParagraphsStructuredDocumentTags.Count == 2);
                Assert.True(document.Sections[1].ParagraphsBookmarks.Count == 1);
                //Assert.True(document.Sections[0].Comments.Count == 0);
                Assert.True(document.Sections[1].ParagraphsImages.Count == 0);
                Assert.True(document.Sections[1].ParagraphsPageBreaks.Count == 0);


                Assert.True(document.Sections[0].Fields.Count == 3);
                Assert.True(document.Sections[0].HyperLinks.Count == 2);

                Assert.True(document.Sections[1].Fields.Count == 0);
                Assert.True(document.Sections[1].HyperLinks.Count == 0);

                Assert.True(document.ParagraphsHyperLinks[0].Hyperlink.IsEmail == false);
                Assert.True(document.ParagraphsHyperLinks[1].Hyperlink.IsEmail == true);
                Assert.True(document.ParagraphsHyperLinks[1].Hyperlink.EmailAddress == "przemyslaw.klys@test.pl");

                Assert.True(document.Sections[0].ParagraphsFields[0].Field.Text == "Przemysław Kłys");
                Assert.True(document.Sections[0].ParagraphsFields[1].Field.Text == "FieldsAndSections.docx");
                Assert.True(document.Sections[0].ParagraphsFields[2].Field.Text == "1");

                Assert.True(document.Sections[0].ParagraphsFields[0].Field.Field == @" AUTHOR  \* Caps  \* MERGEFORMAT ");
                Assert.True(document.Sections[0].ParagraphsFields[1].Field.Field == @" FILENAME   \* MERGEFORMAT ");
                Assert.True(document.Sections[0].ParagraphsFields[2].Field.Field == @" PAGE  \* Arabic  \* MERGEFORMAT ");
            }
        }
        [Fact]
        public void Test_OpeningWordWithImages() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "DocumentWithImages.docx"))) {
                Assert.True(document.Paragraphs.Count == 23);
                Assert.True(document.Sections.Count == 1);
                Assert.True(document.Fields.Count == 0);
                Assert.True(document.HyperLinks.Count == 0);
                Assert.True(document.Equations.Count == 0);
                Assert.True(document.StructuredDocumentTags.Count == 0);
                Assert.True(document.Bookmarks.Count == 0);
                Assert.True(document.Comments.Count == 0);
                Assert.True(document.Images.Count == 7);
            }
        }

    }
}
