using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingCheckBox() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithCheckBox.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var checkBox = document.AddParagraph("Agree:").AddCheckBox(true);

                Assert.Equal(1, document.CheckBoxes.Count);
                Assert.True(checkBox.IsChecked);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(1, document.CheckBoxes.Count);
                Assert.True(document.CheckBoxes[0].IsChecked);

                document.CheckBoxes[0].IsChecked = false;
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.False(document.CheckBoxes[0].IsChecked);
            }
        }
    }
}
