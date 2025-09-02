using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddCustomListWithLevels() {
            var filePath = Path.Combine(_directoryWithFiles, "CustomBulletList.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddCustomList()
                    .AddListLevel(1, WordListLevelKind.BulletSquareSymbol, "Courier New", SixLabors.ImageSharp.Color.Red, fontSize: 14)
                    .AddListLevel(5, WordListLevelKind.BulletBlackCircle, "Arial", colorHex: "#00FF00", fontSize: 10);
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

                var level1Props = list.Numbering.Levels[0]._level.NumberingSymbolRunProperties!;
                Assert.Equal("Courier New", level1Props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("ff0000", level1Props.GetFirstChild<Color>()?.Val);
                Assert.Equal("28", level1Props.GetFirstChild<FontSize>()?.Val);

                var level5Props = list.Numbering.Levels[4]._level.NumberingSymbolRunProperties!;
                Assert.Equal("Arial", level5Props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("00ff00", level5Props.GetFirstChild<Color>()?.Val);
                Assert.Equal("20", level5Props.GetFirstChild<FontSize>()?.Val);
            }
        }

        [Fact]
        public void Test_AddCustomBulletList() {
            var filePath = Path.Combine(_directoryWithFiles, "CustomBulletSimple.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddCustomBulletList(WordListLevelKind.BulletDiamondSymbol, "Wingdings", SixLabors.ImageSharp.Color.Blue, fontSize: 12);
                list.AddItem("Item");
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var list = document.Lists[0];
                Assert.Single(list.Numbering.Levels);
                Assert.Equal("◆", list.Numbering.Levels[0].LevelText);

                var props = list.Numbering.Levels[0]._level.NumberingSymbolRunProperties!;
                Assert.Equal("Wingdings", props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("0000ff", props.GetFirstChild<Color>()?.Val);
                Assert.Equal("24", props.GetFirstChild<FontSize>()?.Val);
            }
        }

        [Fact]
        public void Test_CustomListStartingAtThirdLevel() {
            var filePath = Path.Combine(_directoryWithFiles, "CustomListStartAt3.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddCustomList()
                    .AddListLevel(3, WordListLevelKind.BulletBlackCircle, "Arial");
                list.AddItem("Level3", 2);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var list = document.Lists[0];
                Assert.Equal(3, list.Numbering.Levels.Count);
                Assert.Equal("●", list.Numbering.Levels[0].LevelText);
                Assert.Equal("●", list.Numbering.Levels[1].LevelText);
                Assert.Equal("●", list.Numbering.Levels[2].LevelText);
            }
        }

        [Fact]
        public void Test_CustomList_W15RestartAttributePresent() {
            var filePath = Path.Combine(_directoryWithFiles, "CustomListW15.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddCustomList();
                list.AddListLevel(1, WordListLevelKind.BulletSquareSymbol, "Courier New");
                list.AddItem("Item1");
                document.Save(false);
            }

            using (var wordDoc = WordprocessingDocument.Open(filePath, false)) {
                var numbering = wordDoc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
                Assert.NotNull(numbering.LookupNamespace("w15"));
                var abstractNum = numbering.Elements<AbstractNum>().First();
                var restartAttr = abstractNum.GetAttribute("restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml");
                Assert.Equal("0", restartAttr.Value);
            }
        }
    }
}
