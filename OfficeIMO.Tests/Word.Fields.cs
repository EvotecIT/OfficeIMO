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


                Assert.True(document.Fields[0].FieldFormat == WordFieldFormat.Caps);
                Assert.True(document.Fields[0].FieldType == WordFieldType.Author);
                Assert.True(document.Fields[1].FieldFormat == WordFieldFormat.Mergeformat);
                Assert.True(document.Fields[1].FieldType == WordFieldType.FileName);

                //Assert.True(document.Fields[2].FieldFormat == WordFieldFormat.Arabic);
                Assert.True(document.Fields[2].FieldType == WordFieldType.Page);

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

                document.Fields[0].UpdateField = true;
                document.Fields[1].UpdateField = true;
                Assert.True(document.Fields[0].UpdateField == true);
                Assert.True(document.Fields[1].UpdateField == true);
                document.Fields[2].LockField = true;
                Assert.True(document.Fields[2].LockField == true);
                Assert.True(document.Fields[1].LockField == false);
                Assert.True(document.Fields[0].LockField == false);
                document.Save(false);
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


                document.Save(false);
            }
        }

        [Fact]
        public void Test_CreatingWordWithFields() {
            using (WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "DocumentWithFields.docx"))) {

                document.AddParagraph("This is page number ").AddField(WordFieldType.Page);
                document.AddParagraph("Our title is ").AddField(WordFieldType.Title, WordFieldFormat.Caps);
                document.AddParagraph("Our author is ").AddField(WordFieldType.Author);

                Assert.True(document.Fields[0].FieldFormat == WordFieldFormat.Mergeformat);
                Assert.True(document.Fields[0].FieldType == WordFieldType.Page);
                Assert.True(document.Fields[1].FieldFormat == WordFieldFormat.Caps);
                Assert.True(document.Fields[1].FieldType == WordFieldType.Title);
                Assert.True(document.Fields[2].FieldFormat == WordFieldFormat.Mergeformat);
                Assert.True(document.Fields[2].FieldType == WordFieldType.Author);

                Assert.True(document.Paragraphs.Count == 6);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "DocumentWithFields.docx"))) {
                Assert.True(document.Paragraphs.Count == 6);
                Assert.True(document.Fields[0].FieldFormat == WordFieldFormat.Mergeformat);
                Assert.True(document.Fields[0].FieldType == WordFieldType.Page);
                Assert.True(document.Fields[1].FieldFormat == WordFieldFormat.Caps);
                Assert.True(document.Fields[1].FieldType == WordFieldType.Title);
                Assert.True(document.Fields[2].FieldFormat == WordFieldFormat.Mergeformat);
                Assert.True(document.Fields[2].FieldType == WordFieldType.Author);

                var fieldTypes = (WordFieldType[])Enum.GetValues(typeof(WordFieldType));
                foreach (var fieldType in fieldTypes) {
                    var paragraph = document.AddParagraph("field Type " + fieldType.ToString() + ": ").AddField(fieldType);
                    Assert.True(paragraph.Field.FieldType == fieldType, "FieldType matches");
                }

                Assert.True(document.Paragraphs.Count == 6 + fieldTypes.Length * 2);

                foreach (var fieldType in fieldTypes) {
                    var paragraph = document.AddParagraph("field Type " + fieldType.ToString() + ": ").AddField(fieldType, null, true);
                    Assert.True(paragraph.Field.FieldType == fieldType, "FieldType matches");
                }

                Assert.True(document.Paragraphs.Count == 6 + fieldTypes.Length * 4);

                var fieldTypesFormats = (WordFieldFormat[])Enum.GetValues(typeof(WordFieldFormat));

                foreach (var fieldType in fieldTypes) {
                    foreach (var fieldTypeFormat in fieldTypesFormats) {
                        var paragraph = document.AddParagraph("field Type " + fieldType.ToString() + ": ").AddField(fieldType, fieldTypeFormat);
                        Assert.True(paragraph.Field.FieldType == fieldType, "FieldType matches");
                        Assert.True(paragraph.Field.FieldFormat == fieldTypeFormat, "FieldTypeFormat matches");
                    }
                }

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "DocumentWithFields.docx"))) {
                Assert.True(document.Fields[0].FieldFormat == WordFieldFormat.Mergeformat);
                Assert.True(document.Fields[0].FieldType == WordFieldType.Page);
                Assert.True(document.Fields[1].FieldFormat == WordFieldFormat.Caps);
                Assert.True(document.Fields[1].FieldType == WordFieldType.Title);
                Assert.True(document.Fields[2].FieldFormat == WordFieldFormat.Mergeformat);
                Assert.True(document.Fields[2].FieldType == WordFieldType.Author);

                document.Save();
            }
        }
        [Fact]
        public void Test_CreatingWordFieldsAndSwitches() {
            using (WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "DocumentWithFieldSwitches.docx"))) {

                var instructions = new List<String> { "ANSWER", "\"What is the answer of life and everything?\"" };
                var switches = new List<String> { "\\d \"42\"" };

                var parameters = instructions.Concat(switches).ToList();
                document.AddField(WordFieldType.Ask, parameters: parameters, wordFieldFormat: WordFieldFormat.FirstCap);


                var p = document.AddParagraph(" ");
                p.AddField(WordFieldType.Bibliography);

                document.Save();

                Assert.Equal(2, document.Fields.Count);
                Assert.Equal(WordFieldType.Ask, document.Fields[0].FieldType);
                Assert.Equal(WordFieldFormat.FirstCap, document.Fields[0].FieldFormat);
                Assert.Equal(instructions, document.Fields[0].FieldInstructions);
                Assert.Equal(switches, document.Fields[0].FieldSwitches);
            }
        }

    }
}
