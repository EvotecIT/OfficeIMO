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
        // A div applies its alignment to its own text as well as to its children.
        [InlineData("<div align='center'>Aligned</div>", WordParagraphAlignment.Center)]
        [InlineData("<div style='text-align:right'>Aligned</div>", WordParagraphAlignment.Right)]
        [InlineData("<div align='center'><p>Aligned</p></div>", WordParagraphAlignment.Center)]
        // A paragraph's own declaration, a reset included, outranks container and cell alignment,
        // and its align attribute outranks an alignment the container passed down.
        [InlineData("<div style='text-align:right'><p style='text-align:initial'>Aligned</p></div>", null)]
        [InlineData("<table><tr><td style='text-align:center'><p style='text-align:initial'>Aligned</p></td></tr></table>", null)]
        [InlineData("<div style='text-align:center'><p align='left'>Aligned</p></div>", WordParagraphAlignment.Left)]
        [InlineData("<p style='text-align:-webkit-center'>Aligned</p>", WordParagraphAlignment.Center)]
        [InlineData("<p style='text-align:justify-all'>Aligned</p>", WordParagraphAlignment.Both)]
        [InlineData("<table><tr><td style='text-align:justify'><p>Aligned</p></td></tr></table>", WordParagraphAlignment.Both)]
        [InlineData("<table><tr><td style='text-align:justify'><p style='text-align:left'>Aligned</p></td></tr></table>", WordParagraphAlignment.Left)]
        // inherit and unset take the cell's or container's alignment; initial resets it for the element's own text.
        [InlineData("<table><tr><td style='text-align:center'><p style='text-align:inherit'>Aligned</p></td></tr></table>", WordParagraphAlignment.Center)]
        [InlineData("<table><tr><td style='text-align:right'><p style='text-align:unset'>Aligned</p></td></tr></table>", WordParagraphAlignment.Right)]
        [InlineData("<div style='text-align:right'><div style='text-align:initial'>Aligned</div></div>", null)]
        // An inner container's align attribute outranks alignment inherited from an outer one.
        [InlineData("<div style='text-align:right'><div align='left'>Aligned</div></div>", WordParagraphAlignment.Left)]
        // text-align is inherited by descendants that become their own paragraphs.
        [InlineData("<body style='text-align:center'><p>Aligned</p></body>", WordParagraphAlignment.Center)]
        [InlineData("<div style='text-align:justify'><ul><li>Aligned</li></ul></div>", WordParagraphAlignment.Both)]
        [InlineData("<div align='right'><ol><li>Aligned</li></ol></div>", WordParagraphAlignment.Right)]
        [InlineData("<div style='text-align:center'><ul><li style='text-align:left'>Aligned</li></ul></div>", WordParagraphAlignment.Left)]
        // unset inherits like inherit, on the item and on the containers in between.
        [InlineData("<body style='text-align:center'><ul style='text-align:unset'><li>Aligned</li></ul></body>", WordParagraphAlignment.Center)]
        [InlineData("<body style='text-align:center'><ul><li style='text-align:unset'>Aligned</li></ul></body>", WordParagraphAlignment.Center)]
        // An inner div's align attribute outranks the alignment an outer container passed down to it.
        [InlineData("<div style='text-align:right'><div align='left'><ul><li>Aligned</li></ul></div></div>", WordParagraphAlignment.Left)]
        public void HtmlToWord_MapsParagraphAlignmentSources(string html, WordParagraphAlignment? expected) {
            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new HtmlWordExportOptions());
            WordParagraph paragraph = Assert.Single(document.Paragraphs
                .Concat(document.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs)),
                item => item.Text == "Aligned");

            Assert.Equal(expected, paragraph.ParagraphAlignment);
        }
    }
}
