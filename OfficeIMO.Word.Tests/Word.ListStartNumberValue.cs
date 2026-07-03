using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ListLevelStartNumberingValue() {
            var filePath = Path.Combine(_directoryWithFiles, "ListStartNumbering.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddList(WordListStyle.Numbered);
                list.Numbering.Levels[0].SetStartNumberingValue(5);
                list.AddItem("First");
                list.AddItem("Second");
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(5, document.Lists[0].Numbering.Levels[0].StartNumberingValue);
            }
        }
    }
}
