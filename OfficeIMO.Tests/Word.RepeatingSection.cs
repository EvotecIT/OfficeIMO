using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingRepeatingSection() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithRepeatingSection.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var rs = document.AddParagraph().AddRepeatingSection("Section", "RS", "RSTag");

                Assert.Single(document.RepeatingSections);
                Assert.Equal("RS", rs.Alias);
                Assert.Equal("RSTag", rs.Tag);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.RepeatingSections);
                var control = document.GetRepeatingSectionByAlias("RS");
                Assert.NotNull(control);
                Assert.Equal("RSTag", document.GetRepeatingSectionByTag("RSTag")?.Tag);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.RepeatingSections[0].Remove();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.RepeatingSections);
            }
        }
    }
}
