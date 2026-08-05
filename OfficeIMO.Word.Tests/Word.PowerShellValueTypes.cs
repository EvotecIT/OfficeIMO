using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_OfficeIMOValueTypesAreClrEnumsAndMapToOpenXml() {
            Type[] publicTypes = {
                typeof(WordBreakType),
                typeof(WordHeaderFooterType),
                typeof(WordParagraphAlignment),
                typeof(WordUnderlineStyle),
                typeof(WordTableVerticalAlignment),
                typeof(WordDocumentProtectionType),
                typeof(WordTextDirection),
                typeof(WordHorizontalRelativePosition),
                typeof(WordVerticalRelativePosition)
            };

            Assert.All(publicTypes, type => Assert.True(type.IsEnum, $"{type.Name} must remain a CLR enum."));
            Assert.Equal(BreakValues.Page, WordBreakType.Page.ToOpenXml());
            Assert.Equal(HeaderFooterValues.First, WordHeaderFooterType.First.ToOpenXml());
            Assert.Equal(JustificationValues.Center, WordParagraphAlignment.Center.ToOpenXml());
            Assert.Equal(UnderlineValues.Double, WordUnderlineStyle.Double.ToOpenXml());
            Assert.Equal(TableVerticalAlignmentValues.Bottom, WordTableVerticalAlignment.Bottom.ToOpenXml());
            Assert.Equal(DocumentProtectionValues.ReadOnly, WordDocumentProtectionType.ReadOnly.ToOpenXml());
            Assert.Equal(TextDirectionValues.LefToRightTopToBottom, WordTextDirection.LeftToRightTopToBottom.ToOpenXml());
            Assert.Equal(HorizontalRelativePositionValues.Column, WordHorizontalRelativePosition.Column.ToOpenXml());
            Assert.Equal(VerticalRelativePositionValues.Paragraph, WordVerticalRelativePosition.Paragraph.ToOpenXml());

            AssertCompleteMapping<WordBreakType>(value => value.ToOpenXml());
            AssertCompleteMapping<WordHeaderFooterType>(value => value.ToOpenXml());
            AssertCompleteMapping<WordParagraphAlignment>(value => value.ToOpenXml());
            AssertCompleteMapping<WordUnderlineStyle>(value => value.ToOpenXml());
            AssertCompleteMapping<WordTableVerticalAlignment>(value => value.ToOpenXml());
            AssertCompleteMapping<WordDocumentProtectionType>(value => value.ToOpenXml());
            AssertCompleteMapping<WordTextDirection>(value => value.ToOpenXml());
            AssertCompleteMapping<WordHorizontalRelativePosition>(value => value.ToOpenXml());
            AssertCompleteMapping<WordVerticalRelativePosition>(value => value.ToOpenXml());
        }

        [Fact]
        public void Test_OfficeIMOValueTypeMethodsApplyExpectedOpenXmlValues() {
            string filePath = Path.Combine(_directoryWithFiles, "PowerShellSafeWordValues.docx");

            using WordDocument document = WordDocument.Create(filePath);
            WordParagraph paragraph = document.AddParagraph("Formatted");
            paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
            paragraph.SetUnderline(WordUnderlineStyle.Double);
            paragraph.AddBreak(WordBreakType.Page);

            WordSection section = document.Sections[0];
            Assert.NotNull(section.GetOrCreateHeader(WordHeaderFooterType.First));
            Assert.NotNull(section.GetOrCreateFooter(WordHeaderFooterType.Even));

            WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
            cell.VerticalAlignment = WordTableVerticalAlignment.Bottom;
            cell.TextDirection = WordTextDirection.TopToBottomRightToLeft;

            document.Settings.ProtectionPassword = "TemporaryTestOnly";
            document.Settings.SetProtectionType(WordDocumentProtectionType.ReadOnly);

            WordTextBox textBox = document.AddTextBox("Positioned");
            textBox.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Column;
            textBox.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Paragraph;

            Assert.Equal(WordParagraphAlignment.Center, paragraph.ParagraphAlignment);
            Assert.Equal(WordUnderlineStyle.Double, paragraph.Underline);
            Assert.Contains(
                paragraph._paragraph.Descendants<Break>(),
                item => item.Type?.Value == BreakValues.Page);
            Assert.Equal(WordTableVerticalAlignment.Bottom, cell.VerticalAlignment);
            Assert.Equal(WordTextDirection.TopToBottomRightToLeft, cell.TextDirection);
            cell.TextDirection = null;
            Assert.Null(cell.TextDirection);
            Assert.Equal(WordDocumentProtectionType.ReadOnly, document.Settings.ProtectionType);
            Assert.Equal(WordHorizontalRelativePosition.Column, textBox.HorizontalPositionRelativeFrom);
            Assert.Equal(WordVerticalRelativePosition.Paragraph, textBox.VerticalPositionRelativeFrom);
        }

        [Fact]
        public void Test_FluentBuildersAcceptOfficeIMOValueTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentPowerShellSafeWordValues.docx");

            using WordDocument document = WordDocument.Create(filePath);
            document.AsFluent()
                .Paragraph(paragraph => paragraph
                    .Text("Underlined", text => text.Underline(WordUnderlineStyle.WavyDouble))
                    .Break(WordBreakType.Column))
                .End();

            var mainDocument = document._wordprocessingDocument.MainDocumentPart?.Document;
            Assert.Equal(UnderlineValues.WavyDouble, mainDocument?.Descendants<Underline>().Last().Val?.Value);
            Assert.Equal(BreakValues.Column, mainDocument?.Descendants<Break>().Last().Type?.Value);
        }

        private static void AssertCompleteMapping<T>(Func<T, object> map) where T : struct {
            T[] values = Enum.GetValues(typeof(T)).Cast<T>().ToArray();
            Assert.Equal(values.Length, values.Select(map).Distinct().Count());
        }
    }
}
