using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
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

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) {
                var sdtRun = wordDoc.MainDocumentPart!.Document.Body!.Descendants<SdtRun>().Single();
                var properties = sdtRun.SdtProperties!.Elements<W14.SdtContentCheckBox>().Single();
                var checkedState = properties.Elements<W14.CheckedState>().Single();
                var uncheckedState = properties.Elements<W14.UncheckedState>().Single();

                Assert.Equal(WordCheckBox.CheckedStateValue, checkedState.Val?.Value);
                Assert.Equal(WordCheckBox.UncheckedStateValue, uncheckedState.Val?.Value);

                var symbol = sdtRun.SdtContentRun!.Descendants<Text>().Single().Text;
                Assert.Equal(WordCheckBox.UncheckedSymbol, symbol);
            }
        }

        [Fact]
        public void Test_CheckBoxSymbolUpdatesWithState() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithCheckBoxSymbols.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var checkBox = document.AddParagraph("Accept:").AddCheckBox(false, "Accept", "AcceptTag");
                Assert.False(checkBox.IsChecked);
                document.Save(false);
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) {
                var runText = wordDoc.MainDocumentPart!.Document.Body!.Descendants<SdtRun>()
                    .Single().SdtContentRun!.Descendants<Text>().Single().Text;
                Assert.Equal(WordCheckBox.UncheckedSymbol, runText);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.CheckBoxes.Single().IsChecked = true;
                document.Save(false);
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false)) {
                var runText = wordDoc.MainDocumentPart!.Document.Body!.Descendants<SdtRun>()
                    .Single().SdtContentRun!.Descendants<Text>().Single().Text;
                Assert.Equal(WordCheckBox.CheckedSymbol, runText);
            }
        }
    }
}
