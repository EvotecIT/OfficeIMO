using System.IO;
using System.Reflection;
using OfficeIMO.Word;
using Xunit;
using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddImageFromEmbeddedResource() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentFromResource.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph();
            paragraph.AddImageFromResource(Assembly.GetExecutingAssembly(), "OfficeIMO.Tests.Images.Kulek.jpg", 50, 50);

            Assert.Single(document.Images);
            document.Save(false);
        }
    }
}
