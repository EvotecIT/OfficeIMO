using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ParagraphListStyleDetection() {
            var filePath = Path.Combine(_directoryWithFiles, "ParagraphListStyle.docx");
            using (var document = WordDocument.Create(filePath)) {
                var bullet = document.AddList(WordListStyle.Bulleted);
                bullet.AddItem("Bullet");

                var numbered = document.AddList(WordListStyle.Headings111);
                numbered.AddItem("Numbered");

                var custom = document.AddCustomList();
                custom.AddListLevel(1, WordBulletSymbol.Square, "Courier New", colorHex: "#FF0000");
                custom.AddItem("Custom");

                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                var bulletParagraph = document.Paragraphs.First(p => p.Text == "Bullet");
                var numberedParagraph = document.Paragraphs.First(p => p.Text == "Numbered");
                var customParagraph = document.Paragraphs.First(p => p.Text == "Custom");

                Assert.Equal(WordListStyle.Bulleted, bulletParagraph.GetListStyle());
                Assert.Equal(WordListStyle.Headings111, numberedParagraph.GetListStyle());
                Assert.Equal(WordListStyle.Custom, customParagraph.GetListStyle());
            }
        }
    }
}
