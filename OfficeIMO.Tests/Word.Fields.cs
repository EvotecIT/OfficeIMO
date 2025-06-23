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


                Assert.Equal(new[] { WordFieldFormat.Caps, WordFieldFormat.Mergeformat }, document.Fields[0].FieldFormat);
                Assert.True(document.Fields[0].FieldType == WordFieldType.Author);
                Assert.Equal(new[] { WordFieldFormat.Mergeformat }, document.Fields[1].FieldFormat);
                Assert.True(document.Fields[1].FieldType == WordFieldType.FileName);

                //Assert.Equal(new[] { WordFieldFormat.Arabic, WordFieldFormat.Mergeformat }, document.Fields[2].FieldFormat);
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

                Assert.Equal(new[] { WordFieldFormat.Mergeformat }, document.Fields[0].FieldFormat);
                Assert.True(document.Fields[0].FieldType == WordFieldType.Page);
                Assert.Equal(new[] { WordFieldFormat.Caps, WordFieldFormat.Mergeformat }, document.Fields[1].FieldFormat);
                Assert.True(document.Fields[1].FieldType == WordFieldType.Title);
                Assert.Equal(new[] { WordFieldFormat.Mergeformat }, document.Fields[2].FieldFormat);
                Assert.True(document.Fields[2].FieldType == WordFieldType.Author);

                Assert.True(document.Paragraphs.Count == 6);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "DocumentWithFields.docx"))) {
                Assert.True(document.Paragraphs.Count == 6);
                Assert.Equal(new[] { WordFieldFormat.Mergeformat }, document.Fields[0].FieldFormat);
                Assert.True(document.Fields[0].FieldType == WordFieldType.Page);
                Assert.Equal(new[] { WordFieldFormat.Caps, WordFieldFormat.Mergeformat }, document.Fields[1].FieldFormat);
                Assert.True(document.Fields[1].FieldType == WordFieldType.Title);
                Assert.Equal(new[] { WordFieldFormat.Mergeformat }, document.Fields[2].FieldFormat);
                Assert.True(document.Fields[2].FieldType == WordFieldType.Author);

                var fieldTypes = (WordFieldType[])Enum.GetValues(typeof(WordFieldType));
                foreach (var fieldType in fieldTypes) {
                    var paragraph = document.AddParagraph("field Type " + fieldType.ToString() + ": ").AddField(fieldType);
                    Assert.True(paragraph.Field.FieldType == fieldType, "FieldType matches");
                }

                Assert.True(document.Paragraphs.Count == 6 + fieldTypes.Length * 2);

                foreach (var fieldType in fieldTypes) {
                    var paragraph = document.AddParagraph("field Type " + fieldType.ToString() + ": ").AddField(fieldType, null, advanced: true);
                    Assert.True(paragraph.Field.FieldType == fieldType, "FieldType matches");
                }

                Assert.True(document.Paragraphs.Count == 6 + fieldTypes.Length * 4);

                var fieldTypesFormats = Enum.GetValues(typeof(WordFieldFormat))
                    .Cast<WordFieldFormat>()
                    .GroupBy(f => f.ToString().ToUpperInvariant())
                    .Select(g => g.First())
                    .ToArray();

                foreach (var fieldType in fieldTypes) {
                    foreach (var fieldTypeFormat in fieldTypesFormats) {
                        var paragraph = document.AddParagraph("field Type " + fieldType.ToString() + ": ").AddField(fieldType, fieldTypeFormat);
                        Assert.True(paragraph.Field.FieldType == fieldType, "FieldType matches");
                        Assert.Contains(fieldTypeFormat, paragraph.Field.FieldFormat);
                    }
                }

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "DocumentWithFields.docx"))) {
                Assert.Equal(new[] { WordFieldFormat.Mergeformat }, document.Fields[0].FieldFormat);
                Assert.True(document.Fields[0].FieldType == WordFieldType.Page);
                Assert.Equal(new[] { WordFieldFormat.Caps, WordFieldFormat.Mergeformat }, document.Fields[1].FieldFormat);
                Assert.True(document.Fields[1].FieldType == WordFieldType.Title);
                Assert.Equal(new[] { WordFieldFormat.Mergeformat }, document.Fields[2].FieldFormat);
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
                Assert.Equal(new[] { WordFieldFormat.FirstCap, WordFieldFormat.Mergeformat }, document.Fields[0].FieldFormat);
                Assert.Equal(instructions, document.Fields[0].FieldInstructions);
                Assert.Equal(switches, document.Fields[0].FieldSwitches);
            }
        }

        [Fact]
        public void Test_CreatingWordFieldsTyped() {
            using (WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "DocumentWithFieldSwitchesTyped.docx"))) {
                var ask = new AskField {
                    Bookmark = "ANSWER",
                    Prompt = "What is the answer of life and everything?",
                    DefaultResponse = "42"
                };

                document.AddField(ask, wordFieldFormat: WordFieldFormat.FirstCap);

                document.Save();

                Assert.Single(document.Fields);
                Assert.Equal(WordFieldType.Ask, document.Fields[0].FieldType);
                Assert.Equal(new[] { WordFieldFormat.FirstCap, WordFieldFormat.Mergeformat }, document.Fields[0].FieldFormat);
                Assert.Equal(new[] { "ANSWER", "\"What is the answer of life and everything?\"" }, document.Fields[0].FieldInstructions);
                Assert.Equal(new[] { "\\d \"42\"" }, document.Fields[0].FieldSwitches);
            }
        }

        [Fact]
        public void Test_AdditionalTypedFields() {
            using (WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "DocumentWithAdditionalTypedFields.docx"))) {
                document.AddField(new AuthorField { Author = "John Doe" });
                document.AddField(new FileNameField { IncludePath = true });
                document.AddField(new SetField { Bookmark = "ANSWER", Value = "42" });
                document.AddField(new RefField { Bookmark = "ANSWER", InsertHyperlink = true });

                document.Save();

                Assert.Equal(4, document.Fields.Count);
                Assert.Equal(WordFieldType.Author, document.Fields[0].FieldType);
                Assert.Equal(new[] { "\"John Doe\"" }, document.Fields[0].FieldInstructions);

                Assert.Equal(WordFieldType.FileName, document.Fields[1].FieldType);
                Assert.Equal(new[] { "\\p" }, document.Fields[1].FieldSwitches);

                Assert.Equal(WordFieldType.Set, document.Fields[2].FieldType);
                Assert.Equal(new[] { "ANSWER", "\"42\"" }, document.Fields[2].FieldInstructions);

                Assert.Equal(WordFieldType.Ref, document.Fields[3].FieldType);
                Assert.Equal(new[] { "ANSWER" }, document.Fields[3].FieldInstructions);
                Assert.Equal(new[] { "\\h" }, document.Fields[3].FieldSwitches);
            }
        }

        [Fact]
        public void Test_FieldWithMultipleSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldMultipleSwitches.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph().AddField(WordFieldType.Page, WordFieldFormat.Caps);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(new[] { WordFieldFormat.Caps, WordFieldFormat.Mergeformat }, document.Fields[0].FieldFormat);
            }
        }

        [Fact]
        public void Test_FieldWithCustomFormat() {
            string filePath = Path.Combine(_directoryWithFiles, "CustomFormattedField.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Time: ").AddField(WordFieldType.Time, customFormat: "dd/MM/yyyy HH:mm");
                Assert.Equal(" TIME  \\@ \"dd/MM/yyyy HH:mm\" \\* MERGEFORMAT ", paragraph.Field.Field);
                document.Save(false);
            }

                using (WordDocument document = WordDocument.Load(filePath)) {
                    Assert.Single(document.Fields);
                    Assert.Equal(" TIME  \\@ \"dd/MM/yyyy HH:mm\" \\* MERGEFORMAT ", document.Fields[0].Field);
                    Assert.Equal(WordFieldType.Time, document.Fields[0].FieldType);
                }
        }

        [Fact]
        public void Test_FieldWithNewFormats() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldWithFormats.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var p1 = document.AddParagraph("Page as words: ").AddField(WordFieldType.Page, WordFieldFormat.CardText);
                var p2 = document.AddParagraph("Page ordinal: ").AddField(WordFieldType.Page, WordFieldFormat.Ordinal);
                var p3 = document.AddParagraph("Page hex: ").AddField(WordFieldType.Page, WordFieldFormat.Hex);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(3, document.Fields.Count);
                Assert.Equal(new[] { WordFieldFormat.CardText, WordFieldFormat.Mergeformat }, document.Fields[0].FieldFormat);
                Assert.Equal(new[] { WordFieldFormat.Ordinal, WordFieldFormat.Mergeformat }, document.Fields[1].FieldFormat);
                Assert.Equal(new[] { WordFieldFormat.Hex, WordFieldFormat.Mergeformat }, document.Fields[2].FieldFormat);
            }
        }

        [Fact]
        public void Test_ReadingOfFragmentedInstructions() {
            using WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "partitionedFieldInstructions.docx"));

            Assert.Equal(WordFieldType.XE, document.Fields[0].FieldType);
            Assert.Empty(document.Fields[0].FieldFormat);
            Assert.Equal("\"Introduction\"", document.Fields[0].FieldInstructions.First());

            Assert.Equal(WordFieldType.XE, document.Fields[1].FieldType);
            Assert.Empty(document.Fields[1].FieldFormat);
            Assert.Equal("\"Header 1\"", document.Fields[1].FieldInstructions.First());

            Assert.Equal(WordFieldType.Ask, document.Fields[2].FieldType);
            Assert.Equal(new[] { WordFieldFormat.Mergeformat }, document.Fields[2].FieldFormat);
            Assert.Equal("\"What is the weather today?\"", document.Fields[2].FieldInstructions.First());
            Assert.Equal("\\d \"fine\"", document.Fields[2].FieldSwitches.First());
        }
    }
}
