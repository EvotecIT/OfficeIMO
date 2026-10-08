using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_OmittedTableLayoutRetainsAutofitAcrossNativeDocRoundTrip() {
        using WordDocument source = WordDocument.Create();
        WordTable table = source.AddTable(1, 1, WordTableStyle.TableGrid);
        table.WidthType = WordTableWidthUnit.Auto;
        table.Width = 0;
        table.GridColumnWidth = new List<int> { 2160 };
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.WidthType = WordTableWidthUnit.Dxa;
        cell.Width = 2160;
        cell.Paragraphs[0].Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        cell.Paragraphs[0].FontFamily = "Arial";
        cell.Paragraphs[0].FontSize = 12;
        table._tableProperties!.TableLayout = null;
        string sourceXml = source._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        double originalWidth = TableLayoutContractPdfWidth(source);
        Assert.Equal(sourceXml, source._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        using WordDocument native = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.Equal(WordTableLayoutMode.AutoFit, native.Tables[0].LayoutMode);
        Assert.Equal(originalWidth, TableLayoutContractPdfWidth(native), 2);
        table._tableProperties.TableLayout = new W.TableLayout { Type = W.TableLayoutValues.Autofit };
        Assert.Equal(originalWidth, TableLayoutContractPdfWidth(source), 2);
    }

    private static double TableLayoutContractPdfWidth(WordDocument document) {
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(BorderFramePdfOptions()));
        var page = Assert.Single(pdf.GetPages());
        var boxes = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(box => box.HasValue).Select(box => box!.Value).ToArray();
        Assert.NotEmpty(boxes);
        return boxes.Max(box => box.Right) - boxes.Min(box => box.Left);
    }
}
