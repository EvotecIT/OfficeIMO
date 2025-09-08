using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ConvertBulletToNumbered() {
            var filePath = Path.Combine(_directoryWithFiles, "ConvertBulletToNumbered.docx");
            int indent;
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddList(WordListStyle.Bulleted);
                list.AddItem("One");
                list.AddItem("Two");
                indent = list.Numbering.Levels.First().IndentationLeft;
                list.ConvertToNumbered();
                Assert.Equal(NumberFormatValues.Decimal, list.Numbering!.Levels.First()._level.NumberingFormat!.Val!.Value);
                Assert.Equal(indent, list.Numbering.Levels.First().IndentationLeft);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var paragraph = document.Paragraphs.FirstOrDefault();
                Assert.NotNull(paragraph);
                var list = document.Lists.First();
                Assert.Equal(new[] { "One", "Two" }, list.ListItems.Select(i => i.Text).ToArray());
                Assert.NotNull(list.Numbering);
                Assert.Equal(NumberFormatValues.Decimal, list.Numbering!.Levels.First()._level.NumberingFormat!.Val!.Value);
                Assert.Equal(indent, list.Numbering.Levels.First().IndentationLeft);
            }
        }

        [Fact]
        public void Test_ConvertNumberedToBullet() {
            var filePath = Path.Combine(_directoryWithFiles, "ConvertNumberedToBullet.docx");
            int indent;
            using (var document = WordDocument.Create(filePath)) {
                var list = document.AddList(WordListStyle.Numbered);
                list.AddItem("One");
                list.AddItem("Two");
                indent = list.Numbering.Levels.First().IndentationLeft;
                list.ConvertToBulleted();
                Assert.Equal(NumberFormatValues.Bullet, list.Numbering.Levels.First()._level.NumberingFormat!.Val!.Value);
                Assert.Equal(indent, list.Numbering.Levels.First().IndentationLeft);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var paragraph = document.Paragraphs.FirstOrDefault();
                Assert.NotNull(paragraph);
                var list = document.Lists.First();
                Assert.Equal(new[] { "One", "Two" }, list.ListItems.Select(i => i.Text).ToArray());
                Assert.NotNull(list.Numbering);
                Assert.Equal(NumberFormatValues.Bullet, list.Numbering!.Levels.First()._level.NumberingFormat!.Val!.Value);
                Assert.Equal(indent, list.Numbering.Levels.First().IndentationLeft);
            }
        }
    }
}
