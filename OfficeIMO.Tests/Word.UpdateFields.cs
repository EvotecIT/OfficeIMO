using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_UpdateFieldsUpdatesPageNumbers() {
            string filePath = Path.Combine(_directoryWithFiles, "UpdateFields.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Page 1").AddPageNumber(includeTotalPages: true);
                document.AddPageBreak();
                document.AddParagraph("Page 2");
                document.AddTableOfContent();
                document.UpdateFields();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Contains("1", document.Fields.First(f => f.FieldType == WordFieldType.Page).Text);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_UpdateFieldsWithoutToc() {
            string filePath = Path.Combine(_directoryWithFiles, "UpdateFieldsNoToc.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Page 1").AddPageNumber(includeTotalPages: true);
                document.AddPageBreak();
                document.AddParagraph("Page 2").AddPageNumber();
                document.AddPageBreak();
                document.AddParagraph("Page 3");
                document.UpdateFields();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var pages = document.Fields.Where(f => f.FieldType == WordFieldType.Page).Select(f => f.Text).ToList();
                Assert.Equal(2, pages.Count);
                Assert.Contains("1", pages);
                Assert.Contains("2", pages);
                string total = document.Fields.First(f => f.FieldType == WordFieldType.NumPages).Text;
                Assert.Equal("3", total);
                Assert.False(document.Settings.UpdateFieldsOnOpen);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_UpdateFieldsPreservesSettingWhenTocPresent() {
            string filePath = Path.Combine(_directoryWithFiles, "UpdateFieldsToc.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Page 1").AddPageNumber(includeTotalPages: true);
                document.AddPageBreak();
                document.AddParagraph("Page 2");
                document.AddTableOfContent();
                document.UpdateFields();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }
    }
}
