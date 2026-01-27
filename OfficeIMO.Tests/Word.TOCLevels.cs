using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_TableOfContentsLevelsCanBeConfigured() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTOCLevels.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var toc = document.AddTableOfContent(minLevel: 1, maxLevel: 5);

                document.AddParagraph("Heading 1").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Heading 2").SetStyle(WordParagraphStyles.Heading2);
                document.AddParagraph("Heading 3").SetStyle(WordParagraphStyles.Heading3);
                document.AddParagraph("Heading 4").SetStyle(WordParagraphStyles.Heading4);
                document.AddParagraph("Heading 5").SetStyle(WordParagraphStyles.Heading5);

                var instruction = toc.SdtBlock
                    .Descendants<SimpleField>()
                    .Select(field => field.Instruction?.Value ?? field.Instruction)
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

                Assert.NotNull(instruction);
                Assert.Contains("\\o \"1-5\"", instruction);

                document.Save(false);
            }
        }
    }
}
