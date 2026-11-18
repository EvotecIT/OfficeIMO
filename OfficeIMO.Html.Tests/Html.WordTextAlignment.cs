using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Theory]
        [InlineData("<p align='justify'>Aligned</p>", WordParagraphAlignment.Both)]
        [InlineData("<h2 align='center'>Aligned</h2>", WordParagraphAlignment.Center)]
        [InlineData("<p align='right' style='text-align:left'>Aligned</p>", WordParagraphAlignment.Left)]
        // A CSS reset still outranks the presentational attribute.
        [InlineData("<p align='right' style='text-align:initial'>Aligned</p>", null)]
        [InlineData("<p style='text-align:-webkit-center'>Aligned</p>", WordParagraphAlignment.Center)]
        [InlineData("<p style='text-align:justify-all'>Aligned</p>", WordParagraphAlignment.Both)]
        [InlineData("<table><tr><td style='text-align:justify'><p>Aligned</p></td></tr></table>", WordParagraphAlignment.Both)]
        [InlineData("<table><tr><td style='text-align:justify'><p style='text-align:left'>Aligned</p></td></tr></table>", WordParagraphAlignment.Left)]
        public void HtmlToWord_MapsParagraphAlignmentSources(string html, WordParagraphAlignment? expected) {
            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new HtmlWordExportOptions());
            WordParagraph paragraph = Assert.Single(document.Paragraphs
                .Concat(document.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs)),
                item => item.Text == "Aligned");

            Assert.Equal(expected, paragraph.ParagraphAlignment);
        }
    }
}
