using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void WordToPdfRejectsExcessivelyNestedVisibleRunsWithoutRecursion() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        OpenXmlCompositeElement nested = new W.Hyperlink(new W.Run(new W.Text("visible")));
        for (int i = 0; i < 130; i++) {
            nested = i % 2 == 0
                ? new W.SimpleField(nested) { Instruction = " PAGE " }
                : new W.Hyperlink(nested);
        }
        paragraph._paragraph!.Append(nested);

        Assert.Throws<InvalidDataException>(() => word.ToPdfDocumentResult());
    }

    [Fact]
    public void WordToPdfVisitsNestedTableCellTextBoxesOncePerOwner() {
        using WordDocument word = WordDocument.Create();
        WordTable table = word.AddTable(1, 1, WordTableStyle.TableNormal);
        W.Paragraph nested = new(new W.Run(new W.Text("nested leaf")));
        for (int i = 0; i < 28; i++) {
            var box = new V.TextBox(new W.TextBoxContent(nested));
            var shape = new V.Shape(box) {
                Id = "NestedTableBox" + i,
                Type = "#_x0000_t202",
                Style = "position:absolute;left:0;top:0;width:24pt;height:12pt"
            };
            nested = new W.Paragraph(new W.Run(new W.Picture(shape)));
        }
        table.Rows[0].Cells[0].Paragraphs[0]._paragraph!.Append((W.Run)nested.FirstChild!.CloneNode(true));

        var result = word.ToPdfDocumentResult();
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void WordToPdfDoesNotCountNestedTableCellImageAtEveryTextBoxAncestor() {
        using WordDocument word = WordDocument.Create();
        WordParagraph cell = word.AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0].Paragraphs[0];
        byte[] bytes = OfficeRasterImageEncoder.Encode(
            new OfficeRasterImage(2, 2, OfficeColor.Red), OfficeImageExportFormat.Png);
        using (var imageStream = new MemoryStream(bytes)) cell.InsertImage(imageStream, "nested.png", 10, 10);
        W.Run imageRun = (W.Run)cell._run!.CloneNode(true);
        cell._run.Remove();
        W.Paragraph nested = new(imageRun);
        for (int i = 0; i < 4; i++) {
            var box = new V.TextBox(new W.TextBoxContent(nested));
            var shape = new V.Shape(box) {
                Id = "NestedImageBox" + i,
                Type = "#_x0000_t202",
                Style = "position:absolute;left:0;top:0;width:24pt;height:12pt"
            };
            nested = new W.Paragraph(new W.Run(new W.Picture(shape)));
        }
        cell._paragraph!.Append((W.Run)nested.FirstChild!.CloneNode(true));

        var result = word.ToPdfDocumentResult(new WordToPdfOptions { MaxImagesPerParagraph = 1 });
        Assert.True(result.Succeeded);
        var pdf = PdfCore.PdfDocument.Load(result.Value.ToBytes());
        Assert.Single(pdf.Images.Placements());
    }
}
