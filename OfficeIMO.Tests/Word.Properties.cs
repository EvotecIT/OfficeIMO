using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingDocumentWithProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithProperties.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.BuiltinDocumentProperties.Title = "This is a test for Title";
                document.BuiltinDocumentProperties.Category = "This is a test for Category";

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                Assert.True(document.Paragraphs.Count == 5, "Paragraphs count doesn't match. Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match. Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Paragraphs count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.BuiltinDocumentProperties.Title == "This is a test for Title", "Wrong title");
                Assert.True(document.BuiltinDocumentProperties.Category == "This is a test for Category", "Wrong category");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithProperties.docx"))) {
                Assert.True(document.Paragraphs.Count == 5, "Paragraphs count doesn't match (load). Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match (load). Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Paragraphs count doesn't match for section (load). Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section (load). Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.BuiltinDocumentProperties.Title == "This is a test for Title", "Wrong title (load)");
                Assert.True(document.BuiltinDocumentProperties.Category == "This is a test for Category", "Wrong category (load)");
            }
        }
        [Fact]
        public void Test_CreatingDocumentWithCustomProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithCustomProperties.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.BuiltinDocumentProperties.Title = "This is a test for Title";
                document.BuiltinDocumentProperties.Category = "This is a test for Category";

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                string date = "7/7/2011 10:48";
                DateTime dateTime = DateTime.ParseExact(date, "M/d/yyyy hh:mm", CultureInfo.CurrentCulture);

                document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty { Value = dateTime });
                document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Evotec"));
                document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

                Assert.True(document.ApplicationProperties.Application == "", "Application not matching?");

                document.ApplicationProperties.Application = "OfficeIMO C#";
                document.ApplicationProperties.ApplicationVersion = "1.1.0";

                //Assert.True(document.CustomDocumentProperties["TestProperty"].Value == dateTime, "Custom property should be as expected");
                Assert.True((bool)document.CustomDocumentProperties["IsTodayGreatDay"].Value == true, "Custom property should be as expected");
                Assert.True((string)document.CustomDocumentProperties["MyName"].Value == "Evotec", "Custom property should be as expected");

                Assert.True(document.Paragraphs.Count == 5, "Paragraphs count doesn't match. Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match. Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Paragraphs count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section. Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.BuiltinDocumentProperties.Title == "This is a test for Title", "Wrong title");
                Assert.True(document.BuiltinDocumentProperties.Category == "This is a test for Category", "Wrong category");

                Assert.True(document.ApplicationProperties.Application == "OfficeIMO C#", "Application not matching?");
                Assert.True(document.ApplicationProperties.ApplicationVersion == "1.1.0", "Application version not matching?");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithCustomProperties.docx"))) {
                string date = "7/7/2011 10:48";
                DateTime dateTime = DateTime.ParseExact(date, "M/d/yyyy hh:mm", CultureInfo.CurrentCulture);

                //Assert.True((DateTime) document.CustomDocumentProperties["TestProperty"].Value == dateTime, "Custom property should be as expected");
                Assert.True((bool)document.CustomDocumentProperties["IsTodayGreatDay"].Value == true, "Custom property should be as expected");
                Assert.True((string)document.CustomDocumentProperties["MyName"].Value == "Evotec", "Custom property should be as expected");

                document.CustomDocumentProperties["MyName"].Value = "Przemysław Kłys";

                Assert.True((string)document.CustomDocumentProperties["MyName"].Value == "Przemysław Kłys", "Custom property should be as expected");

                Assert.True(document.Paragraphs.Count == 5, "Paragraphs count doesn't match (load). Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match (load). Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Paragraphs count doesn't match for section (load). Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section (load). Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.BuiltinDocumentProperties.Title == "This is a test for Title", "Wrong title (load)");
                Assert.True(document.BuiltinDocumentProperties.Category == "This is a test for Category", "Wrong category (load)");

                Assert.True(document.ApplicationProperties.Application == "OfficeIMO C#", "Application not matching?");
                Assert.True(document.ApplicationProperties.ApplicationVersion == "1.1.0", "Application version not matching?");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithCustomProperties.docx"))) {

                string date = "7/7/2011 10:48";
                DateTime dateTime = DateTime.ParseExact(date, "M/d/yyyy hh:mm", CultureInfo.CurrentCulture);

                //Assert.True((DateTime)document.CustomDocumentProperties["TestProperty"].Value == dateTime, "Custom property should be as expected");
                Assert.True((bool)document.CustomDocumentProperties["IsTodayGreatDay"].Value == true, "Custom property should be as expected");
                Assert.True((string)document.CustomDocumentProperties["MyName"].Value == "Przemysław Kłys", "Custom property should be as expected");

                Assert.True(document.Paragraphs.Count == 5, "Paragraphs count doesn't match (load). Provided: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count() == 2, "PageBreaks count doesn't match (load). Provided: " + document.PageBreaks.Count);
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Paragraphs count doesn't match for section (load). Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.Sections[0].PageBreaks.Count == 2, "PageBreaks count doesn't match for section (load). Provided: " + document.Sections[0].Paragraphs.Count);
                Assert.True(document.BuiltinDocumentProperties.Title == "This is a test for Title", "Wrong title (load)");
                Assert.True(document.BuiltinDocumentProperties.Category == "This is a test for Category", "Wrong category (load)");

                Assert.True(document.ApplicationProperties.Application == "OfficeIMO C#", "Application not matching?");
                Assert.True(document.ApplicationProperties.ApplicationVersion == "1.1.0", "Application version not matching?");
                document.Save();
            }
        }
        [Fact]
        public void Test_ReadPageBreakProperty() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "EmptyDocumentWithParagraphPropertyPageBreakBefore.docx"))) {
                Assert.True(document.Paragraphs[0].PageBreakBefore);
            }
        }
        [Fact]
        public void Test_SettingPageBreakProperty() {
            bool stateBefore;
            string originalFile = Path.Combine(_directoryDocuments, "EmptyDocumentWithParagraphPropertyPageBreakBefore.docx");
            string tempFile = Path.GetTempFileName();

            using (WordDocument document = WordDocument.Load(originalFile)) {
                stateBefore = document.Paragraphs[0].PageBreakBefore;
                document.Paragraphs[0].PageBreakBefore = false;
                document.Save(tempFile);
            }
            using (WordDocument document = WordDocument.Load(tempFile)) {
                Assert.True(document.Paragraphs[0].PageBreakBefore != stateBefore);
            }
        }
    }
}
