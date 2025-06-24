using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddCustomBulletList() {
            var filePath = Path.Combine(_directoryWithFiles, "CustomBulletList.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddCustomBulletList('\u25A0', "Courier New", "#FF0000", 14);
                list.AddItem("Item 1");
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var list = document.Lists[0];
                Assert.Equal("■", list.Numbering.Levels[0].LevelText);
                Assert.Equal("Courier New", list.FontName);
                Assert.Equal("ff0000", list.ColorHex);
                Assert.Equal(14, list.FontSize);
            }
        }
    }
}
