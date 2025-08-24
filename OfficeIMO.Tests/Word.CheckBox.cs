using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingCheckBox() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithCheckBox.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var checkBox = document.AddParagraph("Agree:").AddCheckBox(true, "Agree", "AgreeTag");

                Assert.Single(document.CheckBoxes);
                Assert.True(checkBox.IsChecked);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.CheckBoxes);
                Assert.True(document.CheckBoxes[0].IsChecked);
                Assert.Equal("AgreeTag", document.CheckBoxes[0].Tag);
                Assert.Equal("Agree", document.CheckBoxes[0].Alias);

                var byTag = document.GetCheckBoxByTag("AgreeTag");
                Assert.NotNull(byTag);
                var byAlias = document.GetCheckBoxByAlias("Agree");
                Assert.NotNull(byAlias);

                byTag.IsChecked = false;
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var checkBox = document.GetCheckBoxByTag("AgreeTag");
                Assert.NotNull(checkBox);
                Assert.False(checkBox!.IsChecked);
            }
        }
    }
}
