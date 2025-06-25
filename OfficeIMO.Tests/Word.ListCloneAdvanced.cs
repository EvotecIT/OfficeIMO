using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CloneListPreservesRestartAfterBreak() {
            var filePath = Path.Combine(_directoryWithFiles, "CloneRestartAfterBreak.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddList(WordListStyle.Bulleted);
                list.RestartNumberingAfterBreak = true;
                list.AddItem("Item 1");
                list.AddItem("Item 2");

                var clone = list.Clone();
                Assert.True(clone.RestartNumberingAfterBreak);

                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.True(document.Lists[0].RestartNumberingAfterBreak);
                Assert.True(document.Lists[1].RestartNumberingAfterBreak);
            }
        }

        [Fact]
        public void Test_CloneListPreservesLevelOverrides() {
            var filePath = Path.Combine(_directoryWithFiles, "CloneLevelOverrides.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddList(WordListStyle.Headings111);
                list.Numbering.Levels[0].SetStartNumberingValue(5);
                list.AddItem("First");
                list.AddItem("Second");

                var clone = list.Clone();
                Assert.Equal(5, clone.Numbering.Levels[0].StartNumberingValue);

                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(5, document.Lists[0].Numbering.Levels[0].StartNumberingValue);
                Assert.Equal(5, document.Lists[1].Numbering.Levels[0].StartNumberingValue);
            }
        }
    }
}
