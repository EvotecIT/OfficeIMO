using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PowerPointInheritedLinkOwnerTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedShapeSlideLinksResolveAndEditTheirActualOwner(bool master) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide first = presentation.AddSlide();
        PowerPointSlide second = presentation.AddSlide();
        PowerPointTextBox box = first.AddTextBoxPoints("Inherited shape", 30, 30, 240, 60);
        box.PlaceholderType = null;
        OpenXmlPart owner = Owner(first, master);
        owner.AddHyperlinkRelationship(PowerPointHyperlinkResolver.CreatePartRelativeUri(owner, second.SlidePart), false, "inherited-link");
        first.SlidePart.AddHyperlinkRelationship(new Uri("https://example.com/wrong-owner"), true, "inherited-link");
        ((P.Shape)box.Element).NonVisualShapeProperties!.NonVisualDrawingProperties!.Append(
            new A.HyperlinkOnClick { Id = "inherited-link", Action = "ppaction://hlinksldjump" });
        box.Element.Remove();
        Tree(owner).Append(box.Element);
        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        PowerPointShape inherited = Assert.Single(reopened.Slides[0].GetInheritedShapesForExport(), shape => shape.Name == box.Name);
        Assert.Equal("#slide-2", inherited.Hyperlink?.OriginalString);
        inherited.SetHyperlink(reopened.Slides[1]);
        Assert.Equal("#slide-2", inherited.Hyperlink?.OriginalString);
        PdfCore.PdfDocumentInfo pdf = PdfCore.PdfInspector.Inspect(reopened.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.LinkAnnotations.Count);
        Assert.All(pdf.LinkAnnotations, link => Assert.Equal("powerpoint-slide-2", link.DestinationName));
    }

    [Theory]
    [InlineData(false, "external")]
    [InlineData(true, "external")]
    [InlineData(false, "slide")]
    [InlineData(true, "slide")]
    [InlineData(false, "next")]
    [InlineData(true, "next")]
    public void InheritedTableCellLinksKeepRelationshipAndActiveSlideContexts(bool master, string kind) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide first = presentation.AddSlide();
        PowerPointSlide second = presentation.AddSlide();
        PowerPointTable table = first.AddTablePoints(1, 1, 30, 30, 240, 60);
        PowerPointTextRun run = table.GetCell(0, 0).AddRun("Inherited link");
        OpenXmlPart owner = Owner(first, master);
        if (kind == "external") owner.AddHyperlinkRelationship(new Uri("https://example.com/correct-owner"), true, "inherited-link");
        else if (kind == "slide") owner.AddHyperlinkRelationship(PowerPointHyperlinkResolver.CreatePartRelativeUri(owner, second.SlidePart), false, "inherited-link");
        first.SlidePart.AddHyperlinkRelationship(new Uri("https://example.com/wrong-owner"), true, "inherited-link");
        (run.Run.RunProperties ??= new A.RunProperties()).Append(new A.HyperlinkOnClick {
            Id = kind == "next" ? string.Empty : "inherited-link",
            Action = kind == "next" ? "ppaction://hlinkshowjump?jump=nextslide" : kind == "slide" ? "ppaction://hlinksldjump" : null
        });
        table.Element.Remove(); Tree(owner).Append(table.Element);
        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        string expected = kind == "external" ? "https://example.com/correct-owner" : "#slide-2";
        PowerPointTable inherited = Assert.Single(reopened.Slides[0].GetInheritedShapesForExport().OfType<PowerPointTable>());
        Assert.Equal(expected, inherited.GetCell(0, 0).Paragraphs[0].Runs.Single(r => r.Text == "Inherited link").Hyperlink?.OriginalString);
        Assert.Equal(expected, inherited.RowItems[0].Cells[0].Runs.Single(r => r.Text == "Inherited link").Hyperlink?.OriginalString);
        Assert.Equal(expected, inherited.ColumnItems[0].GetCell(0).Runs.Single(r => r.Text == "Inherited link").Hyperlink?.OriginalString);
        var conversion = reopened.ConvertToPdfWithEvidence();
        PdfCore.PdfDocumentInfo pdf = PdfCore.PdfInspector.Inspect(conversion.ToBytes());
        Assert.Equal(kind == "next" ? 1 : 2, pdf.LinkAnnotations.Count);
        if (kind == "external") Assert.All(pdf.LinkAnnotations, link => Assert.Equal(expected, link.Uri));
        else Assert.All(pdf.LinkAnnotations, link => Assert.Equal("powerpoint-slide-2", link.DestinationName));
        if (kind == "next") Assert.Contains(conversion.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "internal-slide-target-omitted");
    }

    private static OpenXmlPart Owner(PowerPointSlide slide, bool master) => master
        ? slide.SlidePart.SlideLayoutPart!.SlideMasterPart! : slide.SlidePart.SlideLayoutPart!;

    private static P.ShapeTree Tree(OpenXmlPart owner) => owner is SlideMasterPart master
        ? master.SlideMaster!.CommonSlideData!.ShapeTree! : ((SlideLayoutPart)owner).SlideLayout!.CommonSlideData!.ShapeTree!;
}
