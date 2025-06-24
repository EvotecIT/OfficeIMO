using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddCustomListWithLevels() {
            var filePath = Path.Combine(_directoryWithFiles, "CustomBulletList.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddCustomList()
                    .AddListLevel(1, '\u25A0', "Courier New", "#FF0000", 14)
                    .AddListLevel(5, '\u25CF', "Arial", "#00FF00", 10);
                list.AddItem("Level1");
                list.AddItem("Level5", 4);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var list = document.Lists[0];
                Assert.Equal(5, list.Numbering.Levels.Count);
                Assert.Equal("■", list.Numbering.Levels[0].LevelText);
                Assert.Equal("■", list.Numbering.Levels[3].LevelText);
                Assert.Equal("●", list.Numbering.Levels[4].LevelText);

                var level5Props = list.Numbering.Levels[4]._level.NumberingSymbolRunProperties;
                Assert.Equal("Arial", level5Props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("00ff00", level5Props.GetFirstChild<Color>()?.Val);
                Assert.Equal("20", level5Props.GetFirstChild<FontSize>()?.Val);
            }
        }
    }
}
