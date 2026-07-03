using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ListParser_DetectsTypes() {
            using MemoryStream ms = new MemoryStream();
            using (var document = WordDocument.Create(ms)) {
                var bullet = document.AddList(WordListStyle.Bulleted);
                bullet.AddItem("Bullet 1");
                var ordered = document.AddList(WordListStyle.Numbered);
                ordered.AddItem("Number 1");
                document.Save();
            }
            ms.Position = 0;
            using var word = WordprocessingDocument.Open(ms, false);
            var paragraphs = word.MainDocumentPart!.Document.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
            Assert.True(ListParser.IsBullet(paragraphs[0], word.MainDocumentPart!));
            Assert.True(ListParser.IsOrdered(paragraphs[1], word.MainDocumentPart!));
        }
    }
}
