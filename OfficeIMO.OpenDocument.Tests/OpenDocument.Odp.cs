using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdpTests {
    [Theory]
    [InlineData("libreoffice-impress-basic.odp")]
    [InlineData("microsoft-powerpoint-basic.odp")]
    public void PreservesAuthoredPresentationFixtureOutsideEditedContent(string fixtureName) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", fixtureName);
        using OdpPresentation presentation = OdpPresentation.Open(path);
        var untouched = presentation.Package.Entries
            .Where(entry => entry.Name != "content.xml" && entry.Name != "META-INF/manifest.xml")
            .ToDictionary(entry => entry.Name, entry => entry.GetOriginalBytes());

        presentation.Slides[0].AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "OfficeIMO fixture edit", "OfficeIMOProof");
        byte[] output = presentation.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource });

        using OdpPresentation reopened = OdpPresentation.Open(new MemoryStream(output));
        Assert.Contains(reopened.Slides[0].Shapes.OfType<OdpTextBox>(), textBox =>
            textBox.Paragraphs.Any(paragraph => paragraph.Text == "OfficeIMO fixture edit"));
        foreach (var pair in untouched) Assert.Equal(pair.Value, reopened.Package.GetRequiredEntry(pair.Key).GetOriginalBytes());
    }

    [Fact]
    public void WritesAndReopensSlidesShapesTextImagesTablesAndNotes() {
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        using OdpPresentation presentation = OdpPresentation.Create();
        presentation.PageWidth = OdfLength.Centimeters(33.867);
        presentation.PageHeight = OdfLength.Centimeters(19.05);
        OdpMasterPage master = presentation.AddMasterPage("Brand");
        master.BackgroundColor = OdfColor.Parse("#F7F9FC");
        OdpPresentationLayout layout = presentation.AddLayout("TitleAndContent");
        layout.AddPlaceholder("title", OdfRect.FromCentimeters(2, 1, 29, 2));

        OdpSlide slide = presentation.AddSlide("Overview");
        slide.MasterPageName = master.Name;
        slide.LayoutName = layout.Name;
        slide.BackgroundColor = OdfColor.Parse("#FFFFFF");
        slide.TransitionType = "automatic";
        slide.TransitionStyle = "fade-from-center";
        OdpTextBox title = slide.AddTextBox(OdfRect.FromCentimeters(2, 1, 29, 2), "Native ODP", "Title");
        title.Paragraphs.Single().FontSize = OdfLength.Points(28);
        OdpTextBox content = slide.AddTextBox(OdfRect.FromCentimeters(2, 4, 14, 8), name: "Content");
        OdpParagraph paragraph = content.AddParagraph("Dependency-free ");
        OdpRun run = paragraph.AddRun("presentation");
        run.Bold = true;
        run.Color = OdfColor.Parse("#175CD3");
        OdpList list = content.AddList();
        list.AddItem("Text and lists");
        list.AddItem("Shapes and images");

        OdpRectangle rectangle = slide.AddRectangle(OdfRect.FromCentimeters(18, 4, 5, 3));
        rectangle.FillColor = OdfColor.Parse("#D1E9FF");
        rectangle.StrokeColor = OdfColor.Parse("#175CD3");
        slide.AddEllipse(OdfRect.FromCentimeters(24, 4, 3, 3)).FillColor = OdfColor.Parse("#DCFAE6");
        slide.AddLine(OdfLength.Centimeters(18), OdfLength.Centimeters(8), OdfLength.Centimeters(28), OdfLength.Centimeters(8));
        OdpGroup group = slide.AddGroup("Grouped");
        group.Transform = "translate(1cm 1cm)";
        group.AddRectangle(OdfRect.FromCentimeters(18, 9, 2, 1));
        group.AddEllipse(OdfRect.FromCentimeters(21, 9, 1, 1));
        OdpImage image = slide.AddImage(png, "pixel.png", OdfRect.FromCentimeters(28, 4, 2, 2));
        image.Crop = OdfInsets.FromCentimeters(0, 0, 0, 0);
        OdpTable table = slide.AddTable(OdfRect.FromCentimeters(18, 11, 12, 4), 2, 2, "Metrics");
        table.Cell(0, 0).Text = "Metric";
        table.Cell(0, 1).Text = "Value";
        table.Merge(1, 0, 1, 2).Text = "42";
        slide.GetOrCreateSpeakerNotes().AddParagraph("Explain the native package boundary.");

        OdpSlide hidden = presentation.AddSlide("Appendix");
        hidden.Hidden = true;
        hidden.AddTextBox(OdfRect.FromCentimeters(2, 2, 20, 3), "Hidden appendix");
        presentation.MoveSlide(1, 0);

        byte[] bytes = presentation.ToBytes();
        Assert.True(presentation.Validate().IsValid);
        using OdpPresentation reopened = OdpPresentation.Open(new MemoryStream(bytes));

        Assert.Equal(2, reopened.Slides.Count);
        Assert.Equal("Appendix", reopened.Slides[0].Name);
        Assert.True(reopened.Slides[0].Hidden);
        OdpSlide actual = reopened.Slides[1];
        Assert.Equal("Overview", actual.Name);
        Assert.Equal("Brand", actual.MasterPageName);
        Assert.Equal("fade-from-center", actual.TransitionStyle);
        Assert.Equal(2, actual.Shapes.OfType<OdpTextBox>().Count());
        Assert.Single(actual.Shapes.OfType<OdpImage>());
        Assert.Single(actual.Shapes.OfType<OdpTable>());
        Assert.Equal("42", actual.Shapes.OfType<OdpTable>().Single().Cell(1, 0).Text);
        Assert.True(actual.Shapes.OfType<OdpTable>().Single().Cell(1, 1).IsCovered);
        Assert.Equal("Explain the native package boundary.", actual.SpeakerNotes!.Paragraphs.Single().Text);
        Assert.Contains("presentation-transitions", reopened.InspectFeatures().Findings.Select(finding => finding.Name));
    }
}
