using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ListHelpers() {
            var filePath = Path.Combine(_directoryWithFiles, "ListHelpers.docx");
            using (var document = WordDocument.Create(filePath)) {
                var bullet = document.CreateBulletList();
                bullet.AddItem("One");

                var numbered = document.CreateNumberedList();
                numbered.AddItem("First");
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Lists.Count);
                Assert.Equal("One", document.Lists[0].ListItems[0].Text);
                Assert.Equal("First", document.Lists[1].ListItems[0].Text);
            }
        }
    }
}
