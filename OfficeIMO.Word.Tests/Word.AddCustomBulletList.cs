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
                var list = document.AddCustomList();
                WordListLevel first = list.AddListLevel(1, WordListLevelKind.BulletSquareSymbol, "Courier New", OfficeIMO.OfficeColor.Red, fontSize: global::OfficeIMO.OfficeLength.Points(14));
                WordListLevel fifth = list.AddListLevel(5, WordListLevelKind.BulletBlackCircle, "Arial", colorHex: "#00FF00", fontSize: global::OfficeIMO.OfficeLength.Points(10));
                first.IndentationLeft = OfficeIMO.OfficeLength.Points(36);
                Assert.Throws<OverflowException>(() => first.IndentationLeft = OfficeIMO.OfficeLength.Emus(long.MaxValue));
                Assert.Equal(OfficeIMO.OfficeLength.Points(36), first.IndentationLeft);
                fifth.IndentationHanging = OfficeIMO.OfficeLength.Points(18);
                Assert.Same(first.OpenXmlElement, list.Numbering.Levels[0].OpenXmlElement);
                Assert.Same(fifth.OpenXmlElement, list.Numbering.Levels[4].OpenXmlElement);
                list.AddItem("Level1");
                list.AddItem("Level5", 4);
                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                var list = document.Lists[0];
                Assert.Equal(5, list.Numbering.Levels.Count);
                Assert.Equal("720", list.Numbering.Levels[0].OpenXmlElement.GetFirstChild<PreviousParagraphProperties>()!.Indentation!.Left!.Value);
                Assert.Equal("360", list.Numbering.Levels[4].OpenXmlElement.GetFirstChild<PreviousParagraphProperties>()!.Indentation!.Hanging!.Value);
                Assert.Equal("■", list.Numbering.Levels[0].LevelText);
                Assert.Equal("■", list.Numbering.Levels[3].LevelText);
                Assert.Equal("●", list.Numbering.Levels[4].LevelText);

                var level1Props = list.Numbering.Levels[0].OpenXmlElement.NumberingSymbolRunProperties!;
                Assert.Equal("Courier New", level1Props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("FF0000", level1Props.GetFirstChild<Color>()?.Val);
                Assert.Equal("28", level1Props.GetFirstChild<FontSize>()?.Val);

                var level5Props = list.Numbering.Levels[4].OpenXmlElement.NumberingSymbolRunProperties!;
                Assert.Equal("Arial", level5Props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("00FF00", level5Props.GetFirstChild<Color>()?.Val);
                Assert.Equal("20", level5Props.GetFirstChild<FontSize>()?.Val);
            }
        }

        [Fact]
        public void Test_AddCustomBulletList() {
            var filePath = Path.Combine(_directoryWithFiles, "CustomBulletSimple.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddCustomBulletList(WordListLevelKind.BulletDiamondSymbol, "Wingdings", OfficeIMO.OfficeColor.Blue, fontSize: global::OfficeIMO.OfficeLength.Points(12));
                list.AddItem("Item");
                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                var list = document.Lists[0];
                Assert.Single(list.Numbering.Levels);
                Assert.Equal("◆", list.Numbering.Levels[0].LevelText);

                var props = list.Numbering.Levels[0].OpenXmlElement.NumberingSymbolRunProperties!;
                Assert.Equal("Wingdings", props.GetFirstChild<RunFonts>()?.Ascii);
                Assert.Equal("0000FF", props.GetFirstChild<Color>()?.Val);
                Assert.Equal("24", props.GetFirstChild<FontSize>()?.Val);
            }
        }

        [Fact]
        public void Test_CustomListStartingAtThirdLevel() {
            var filePath = Path.Combine(_directoryWithFiles, "CustomListStartAt3.docx");
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddCustomList();
                WordListLevel third = list.AddListLevel(3, WordListLevelKind.BulletBlackCircle, "Arial");
                Assert.Equal(2, third.LevelIndex);
                list.AddItem("Level3", 2);
                document.Save();
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
                document.Save();
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
