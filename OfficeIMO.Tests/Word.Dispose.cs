using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains tests for disposing Word documents.
    /// </summary>
    public partial class Word {
        [Fact]
        public void Test_DisposeMultipleTimes() {
            var filePath = Path.Combine(_directoryWithFiles, "DisposeTestingMultipleTimes.docx");
            File.Delete(filePath);

            var document = WordDocument.Create(filePath);
            document.AddParagraph("This is my test");
            document.Save();
            document.Dispose();
            document.Dispose();

            Assert.False(filePath.IsFileLocked());
        }
    }
}
