using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FindAndReplace_CaseSensitive() {
            string filePath = Path.Combine(_directoryWithFiles, "CaseSensitiveReplace.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hello World");
                document.AddParagraph("hello world");

                var replacedCount = document.FindAndReplace("Hello World", "Bye", StringComparison.Ordinal);
                Assert.Equal(1, replacedCount);
                Assert.Equal("Bye", document.Paragraphs[0].Text);
                Assert.Equal("hello world", document.Paragraphs[1].Text);

                replacedCount = document.FindAndReplace("hello world", "Bye", StringComparison.OrdinalIgnoreCase);
                Assert.Equal(1, replacedCount);

                Assert.Equal("Bye", document.Paragraphs[1].Text);
                document.Save(false);
            }
        }
    }
}
