using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains combo box content control tests.
    /// </summary>
    public partial class Word {
        [Fact]
        public void Test_AddingComboBox() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithComboBox.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var items = new List<string> { "One", "Two" };
                var cb = document.AddParagraph("Select:").AddComboBox(items, "CB", "CBTag", defaultValue: "Two");

                Assert.Single(document.ComboBoxes);
                Assert.Equal(2, cb.Items.Count);
                Assert.Equal("CB", cb.Alias);
                Assert.Equal("CBTag", cb.Tag);
                Assert.Equal("Two", cb.SelectedValue);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.ComboBoxes);
                var list = document.GetComboBoxByAlias("CB");
                Assert.NotNull(list);
                Assert.Equal("CBTag", document.GetComboBoxByTag("CBTag")?.Tag);
                Assert.Equal("Two", list!.SelectedValue);
                document.Save(false);
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) {
                var runText = wordDoc.MainDocumentPart!.Document.Body!.Descendants<SdtRun>()
                    .Single().SdtContentRun!.Descendants<Text>().SingleOrDefault()?.Text;
                Assert.Equal("Two", runText);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.ComboBoxes[0].Remove();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ComboBoxes);
            }
        }

        [Fact]
        public void Test_ComboBoxDefaultsToFirstItemWhenNoneSpecified() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithComboBoxDefault.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var items = new List<string> { "Alpha", "Beta", "Gamma" };
                var combo = document.AddParagraph("Choose:").AddComboBox(items, "Combo", "ComboTag");
                Assert.Equal("Alpha", combo.SelectedValue);
                document.Save(false);
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) {
                var runText = wordDoc.MainDocumentPart!.Document.Body!.Descendants<SdtRun>()
                    .Single().SdtContentRun!.Descendants<Text>().SingleOrDefault()?.Text;
                Assert.Equal("Alpha", runText);
            }
        }

        [Fact]
        public void Test_ComboBoxThrowsWhenDefaultMissingFromItems() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithComboBoxInvalidDefault.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var items = new List<string> { "Red", "Green" };
                var paragraph = document.AddParagraph("Pick:");
                Assert.Throws<ArgumentException>(() => paragraph.AddComboBox(items, "Combo", "ComboTag", defaultValue: "Blue"));
            }
        }
    }
}
