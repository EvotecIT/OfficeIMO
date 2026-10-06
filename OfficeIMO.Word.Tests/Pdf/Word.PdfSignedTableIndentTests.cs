using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_PreservesSignedTableIndentInBodyAndColumns(bool nativeDoc, bool columns) {
        using WordDocument source = WordDocument.Create();
        if (columns) source.Sections[0].ColumnCount = 2;
        WordTable baseline = source.AddTable(1, 1);
        Configure(baseline, "MarkerDefault");
        source.AddParagraph("Gap");
        WordTable indented = source.AddTable(1, 1);
        Configure(indented, "MarkerNegative");
        indented.StyleDetails!.TableIndentationWidth = -108;
        using WordDocument reopened = WordDocument.Load(new MemoryStream(nativeDoc
            ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        Assert.Equal((short)-108, reopened.Tables[1].StyleDetails!.TableIndentationWidth);
        byte[] bytes = reopened.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PageSize(500, 400), Margins = PageMargins.Uniform(40)
        });
        using var pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToArray();
        var original = Assert.Single(words, word => word.Text == "MarkerDefault");
        var shifted = Assert.Single(words, word => word.Text == "MarkerNegative");
        Assert.Equal(-5.4, shifted.BoundingBox.Left - original.BoundingBox.Left, 2);

        static void Configure(WordTable table, string marker) {
            table.LayoutMode = WordTableLayoutMode.Fixed;
            table.WidthType = WordTableWidthUnit.Dxa;
            table.Width = 3600;
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.WidthType = WordTableWidthUnit.Dxa;
            cell.Width = 3600;
            cell.MarginLeftWidth = 0; cell.MarginRightWidth = 0;
            cell.Paragraphs[0].Text = marker;
        }
    }
}
