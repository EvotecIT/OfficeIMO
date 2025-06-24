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
                    .AddListLevel(1, WordBulletSymbol.Square, "Courier New", SixLabors.ImageSharp.Color.Red, fontSize: 14)
                    .AddListLevel(5, WordBulletSymbol.BlackCircle, "Arial", colorHex: "#00FF00", fontSize: 10);
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

                var level1Props = list.Numbering.Levels[0]._level.NumberingSymbolRunProperties;
                Assert.Equal("Courier New", level1Props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("ff0000", level1Props.GetFirstChild<Color>()?.Val);
                Assert.Equal("28", level1Props.GetFirstChild<FontSize>()?.Val);

                var level5Props = list.Numbering.Levels[4]._level.NumberingSymbolRunProperties;
                Assert.Equal("Arial", level5Props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("00ff00", level5Props.GetFirstChild<Color>()?.Val);
                Assert.Equal("20", level5Props.GetFirstChild<FontSize>()?.Val);
            }
        }

        [Fact]
        public void Test_AddCustomBulletList() {
            var filePath = Path.Combine(_directoryWithFiles, "CustomBulletSimple.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddCustomBulletList(WordBulletSymbol.Diamond, "Wingdings", SixLabors.ImageSharp.Color.Blue, fontSize: 12);
                list.AddItem("Item");
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var list = document.Lists[0];
                Assert.Single(list.Numbering.Levels);
                Assert.Equal("◆", list.Numbering.Levels[0].LevelText);

                var props = list.Numbering.Levels[0]._level.NumberingSymbolRunProperties;
                Assert.Equal("Wingdings", props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("0000ff", props.GetFirstChild<Color>()?.Val);
                Assert.Equal("24", props.GetFirstChild<FontSize>()?.Val);
            }
        }
    }
}
