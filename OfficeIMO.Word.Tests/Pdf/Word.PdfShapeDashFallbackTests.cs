using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfShapeDashFallbackTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingPresetDashValueRetainsSolidShapeAndText(bool otherNamespace) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Before");
        document.AddParagraph().AddShapeDrawing(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(96), OfficeIMO.OfficeLength.Points(44));
        document.AddParagraph("After");
        var shape = document._wordprocessingDocument.MainDocumentPart!.Document.Descendants<Wps.ShapeProperties>().First();
        var dash = new A.PresetDash();
        if (otherNamespace) dash.SetAttribute(new OpenXmlAttribute("x", "val", "urn:other", "dash"));
        shape.GetFirstChild<A.Outline>()?.Remove();
        shape.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }), dash) { Width = 25400 });
        using var loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using var pdf = PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal("BeforeAfter", string.Concat(pdf.GetPage(1).Letters.Select(x => x.Value)));
    }
}
