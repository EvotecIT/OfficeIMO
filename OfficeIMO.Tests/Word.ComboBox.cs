using System.Collections.Generic;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingComboBox() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithComboBox.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var items = new List<string> { "One", "Two" };
                var cb = document.AddParagraph("Select:").AddComboBox(items, "CB", "CBTag");

                Assert.Single(document.ComboBoxes);
                Assert.Equal(2, cb.Items.Count);
                Assert.Equal("CB", cb.Alias);
                Assert.Equal("CBTag", cb.Tag);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.ComboBoxes);
                var list = document.GetComboBoxByAlias("CB");
                Assert.NotNull(list);
                Assert.Equal("CBTag", document.GetComboBoxByTag("CBTag")?.Tag);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.ComboBoxes[0].Remove();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ComboBoxes);
            }
        }
    }
}
