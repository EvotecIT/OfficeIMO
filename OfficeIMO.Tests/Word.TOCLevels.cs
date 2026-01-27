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

        [Fact]
        public void Test_TableOfContentsLevelsUpdateComplexFields() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTOCLevelsComplex.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var toc = document.AddTableOfContent();

                var simpleField = toc.SdtBlock.Descendants<SimpleField>().First();
                var instructionText = simpleField.Instruction?.Value ?? simpleField.Instruction;
                instructionText ??= " TOC \\o \"1-3\" \\h \\z \\u ";
                var paragraph = simpleField.Ancestors<Paragraph>().First();

                // Simulate Word converting the TOC into a complex field.
                simpleField.Remove();
                paragraph.Append(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode { Text = instructionText }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("No table of contents entries found.")),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

                toc.SetLevels(1, 5);

                var fieldCodeText = toc.SdtBlock
                    .Descendants<FieldCode>()
                    .Select(code => code.Text)
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

                Assert.NotNull(fieldCodeText);
                Assert.Contains("\\o \"1-5\"", fieldCodeText);

                document.Save(false);
            }
        }
    }
}
