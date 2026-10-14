using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public class ReaderDrawingFormatTests {
    [Theory]
    [InlineData("odg")]
    [InlineData("fodg")]
    public void DrawingHandlerExtractsNestedTextAndEnforcesAggregateBudget(string extension) {
        OdgDocument drawing = OdgDocument.Create();
        OdgPage page = drawing.AddPage("Drawing");
        page.Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 4, 2), "First");
        OdgShape group = page.Shapes.AddGroup();
        group.Children.AddRectangle(OdfRect.FromCentimeters(1, 3, 4, 2)).Text = "Nested";
        byte[] image = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Blue));
        group.Children.AddImage(image, "caption.png", OdfRect.FromCentimeters(6, 3, 4, 2)).AddParagraph("Caption");
        using var stream = new MemoryStream();
        if (extension == "fodg") drawing.SaveFlatXml(stream); else drawing.Save(stream);
        var reader = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build();
        ReaderChunk chunk = Assert.Single(reader.Read(stream.ToArray(), "drawing." + extension));
        Assert.Equal("First" + Environment.NewLine + "Nested" + Environment.NewLine + "Caption", chunk.Text);
        Assert.Equal(1, chunk.Location.Page);
        Assert.Contains(chunk.Warnings!, warning => warning.Contains("hidden layers"));
        // The page name and the two original labels fit; the image caption must also charge the budget.
        var bounded = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler(new ReaderOpenDocumentOptions { MaxExtractedCharacters = 18 }).Build();
        Assert.Throws<InvalidDataException>(() => bounded.Read(stream.ToArray(), "drawing." + extension).ToArray());
    }

    [Theory]
    [InlineData("vdx", VisioPackageType.Drawing)]
    [InlineData("vtx", VisioPackageType.Template)]
    [InlineData("vsx", VisioPackageType.Stencil)]
    public void LegacyHandlerDispatchesFamiliesAndCarriesImportDiagnostics(string extension, VisioPackageType family) {
        VisioDocument drawing = VisioDocument.Create(family);
        if (family == VisioPackageType.Stencil) drawing.RegisterMaster("Box", new VisioShape("1", 1, 1, 2, 1, "Reusable box"));
        else drawing.AddPage("Workflow").Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Start workflow"));
        byte[] bytes = drawing.ToLegacyXmlResult().Value;
        var reader = new OfficeDocumentReaderBuilder().AddVisioHandler().Build();
        using var stream = new MemoryStream(bytes);
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "workflow." + extension);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "VDX_MODEL_PROFILE");
        if (family != VisioPackageType.Stencil) {
            ReaderChunk chunk = Assert.Single(reader.Read(bytes, "workflow." + extension));
            Assert.Contains("Start workflow", chunk.Text);
            Assert.Contains(chunk.Warnings!, warning => warning.Contains("existing Visio model"));
        }
        Assert.True(stream.CanRead);
    }
}
