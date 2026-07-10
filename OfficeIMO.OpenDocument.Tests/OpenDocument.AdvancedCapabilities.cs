using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentAdvancedCapabilityTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public void AuthorsInspectsAcceptsAndRejectsTrackedParagraphChanges() {
        using OdtDocument document = OdtDocument.Create();
        OdtParagraph removed = document.AddParagraph("Restore me");
        OdtTrackedChange deletion = document.DeleteParagraphTracked(removed, "Reviewer", new DateTimeOffset(2026, 7, 10, 8, 0, 0, TimeSpan.Zero));
        OdtTrackedChange insertion = document.AddTrackedParagraphInsertion("Keep me", "Author", new DateTimeOffset(2026, 7, 10, 9, 0, 0, TimeSpan.Zero));

        Assert.Equal(2, document.TrackedChanges.Count);
        Assert.Equal("Restore me", deletion.DeletedText);
        Assert.Single(document.ContentBlocks);
        Assert.Equal("Keep me", document.ContentBlocks[0].Paragraph!.Text);
        Assert.Contains(document.InspectFeatures().Findings, finding => finding.Name == "tracked-changes");

        Assert.True(document.RejectTrackedChange(deletion.Id));
        Assert.True(document.AcceptTrackedChange(insertion.Id));
        Assert.Empty(document.TrackedChanges);
        Assert.Equal(new[] { "Restore me", "Keep me" }, document.ContentBlocks.Select(block => block.Paragraph!.Text));

        using OdtDocument reopened = OdtDocument.Open(new MemoryStream(document.ToBytes()));
        Assert.Equal(new[] { "Restore me", "Keep me" }, reopened.ContentBlocks.Select(block => block.Paragraph!.Text));
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void AuthorsAndRoundTripsBasicPresentationAnimation() {
        using OdpPresentation presentation = OdpPresentation.Create();
        OdpSlide slide = presentation.AddSlide("Animated");
        OdpRectangle shape = slide.AddRectangle(OdfRect.FromCentimeters(2, 2, 5, 3));
        OdpAnimation animation = slide.AddFadeInAnimation(shape, TimeSpan.FromSeconds(1.5));
        OdpEllipse laterShape = slide.AddEllipse(OdfRect.FromCentimeters(8, 2, 3, 3));

        Assert.NotNull(shape.XmlId);
        Assert.Throws<ArgumentException>(() => laterShape.XmlId = shape.XmlId);
        Assert.Equal(shape.XmlId, animation.TargetElement);
        Assert.Equal("opacity", animation.AttributeName);
        Assert.Equal(TimeSpan.FromSeconds(1.5), animation.Duration);
        Assert.Contains(presentation.InspectFeatures().Findings,
            finding => finding.Name == "presentation-animations" && finding.Support == OdfFeatureSupport.Editable);

        using OdpPresentation reopened = OdpPresentation.Open(new MemoryStream(presentation.ToBytes()));
        OdpAnimation actual = Assert.Single(Assert.Single(reopened.Slides).Animations);
        Assert.Equal("0", actual.From);
        Assert.Equal("1", actual.To);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void TrackedParagraphDeletionRejectsNestedTableContent() {
        using OdtDocument document = OdtDocument.Create();
        OdtParagraph nested = document.AddTable(1, 1, "Nested").Cell(0, 0).Paragraphs[0];

        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            document.DeleteParagraphTracked(nested, "Reviewer"));

        Assert.Contains("top-level", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(document.TrackedChanges);
        Assert.True(document.Validate().IsValid);
    }

    [Fact]
    public void ProjectsPackageToFlatXmlAndBackWithEmbeddedImageBytes() {
        using OdtDocument document = OdtDocument.Create();
        OdtParagraph paragraph = document.AddParagraph("Flat XML");
        paragraph.AddImage(TinyPng, "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        XDocument flat = document.ToFlatXml();
        Assert.Equal(OdfMediaTypes.Text, (string?)flat.Root!.Attribute(XName.Get("mimetype", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")));
        Assert.Contains(flat.Descendants(), element => element.Name.LocalName == "binary-data");

        using var stream = new MemoryStream();
        document.SaveFlatXml(stream);
        stream.Position = 0;
        using OdtDocument reopened = OdtDocument.OpenFlatXml(stream);
        Assert.Equal("Flat XML", reopened.ContentBlocks[0].Paragraph!.Text);
        Assert.Equal(TinyPng, Assert.Single(reopened.ContentBlocks[0].Paragraph!.Images).GetImageBytes());
        OdfValidationResult validation = reopened.Validate();
        Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics.Select(item => item.Id + ": " + item.Message)));
    }

    [Fact]
    public void FlatXmlRoundTripsHeaderImagesAndReportsThemAsRepresented() {
        using OdtDocument document = OdtDocument.Create();
        OdtImage image = document.PageLayout.Header.AddParagraph("Logo").AddImage(TinyPng, "header.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        using var stream = new MemoryStream();

        document.SaveFlatXml(stream);

        Assert.DoesNotContain(image.Path, document.LastSaveReport!.LossyEntries);
        stream.Position = 0;
        using OdtDocument reopened = OdtDocument.OpenFlatXml(stream);
        Assert.Equal(TinyPng, reopened.PageLayout.Header.Paragraphs.Single().Images.Single().GetImageBytes());
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void FlatXmlExportToleratesMissingOptionalSettingsPart() {
        using OdtDocument document = OdtDocument.Create();
        document.AddParagraph("No settings");
        document.Package.RemoveEntry("settings.xml");
        using var stream = new MemoryStream();

        document.SaveFlatXml(stream);

        stream.Position = 0;
        using OdtDocument reopened = OdtDocument.OpenFlatXml(stream);
        Assert.Equal("No settings", reopened.Paragraphs.Single().Text);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void FlatXmlRoundTripsAllThreeDocumentKinds() {
        using OdtDocument text = OdtDocument.Create();
        text.AddParagraph("Text");
        using OdsDocument sheet = OdsDocument.Create();
        sheet.AddSheet("Data").Cell(0, 0).SetNumber(7D);
        using OdpPresentation slides = OdpPresentation.Create();
        slides.AddSlide("One").AddTextBox(OdfRect.FromCentimeters(1, 1, 5, 2), "Slide");

        using OdfDocument reopenedText = ReopenFlat(text);
        using OdfDocument reopenedSheet = ReopenFlat(sheet);
        using OdfDocument reopenedSlides = ReopenFlat(slides);

        Assert.IsType<OdtDocument>(reopenedText);
        Assert.IsType<OdsDocument>(reopenedSheet);
        Assert.IsType<OdpPresentation>(reopenedSlides);
        Assert.True(reopenedText.Validate().IsValid);
        Assert.True(reopenedSheet.Validate().IsValid);
        Assert.True(reopenedSlides.Validate().IsValid);
    }

    [Fact]
    public void FlatXmlRestoresStylesScopedAutomaticStylesAndSourceVersion() {
        using OdtDocument document = OdtDocument.Create();
        document.PageLayout.MarginLeft = OdfLength.Centimeters(3.25);
        OdtParagraph header = document.PageLayout.Header.AddParagraph("Styled header");
        header.Bold = true;
        XDocument flat = document.ToFlatXml();
        flat.Root!.SetAttributeValue(OdfNamespaces.Office + "version", "1.3");
        using var stream = new MemoryStream();
        flat.Save(stream, SaveOptions.DisableFormatting);
        stream.Position = 0;

        using OdtDocument reopened = OdtDocument.OpenFlatXml(stream);

        Assert.Equal(OdfVersion.V1_3, reopened.Version);
        Assert.Equal(OdfLength.Centimeters(3.25), reopened.PageLayout.MarginLeft);
        Assert.True(reopened.PageLayout.Header.Paragraphs.Single().Bold);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void FlatXmlSaveReportsPackageOnlyContentAsLossy() {
        using OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Flat projection");
        document.Package.AddOrReplaceEntry("Thumbnails/thumbnail.png", TinyPng, "image/png");
        using var output = new MemoryStream();

        document.SaveFlatXml(output);

        Assert.NotNull(document.LastSaveReport);
        Assert.Contains("Thumbnails/thumbnail.png", document.LastSaveReport!.LossyEntries);
        Assert.Contains("content.xml", document.LastSaveReport.RewrittenEntries);
    }

    [Fact]
    public void AdvancedCapabilityLinesAreStableAndDistinct() {
        string[] expected = { "formula-evaluation", "tracked-change-editing", "advanced-charts", "presentation-animations", "flat-xml", "encryption", "digital-signatures" };
        Assert.Equal(expected, OdfCapabilityCatalog.Advanced.Select(capability => capability.Id));
        Assert.Equal(OdfCapabilityLevel.DetectedUnsupported, OdfCapabilityCatalog.Find("encryption")!.Level);
        Assert.Equal(OdfCapabilityLevel.Preserved, OdfCapabilityCatalog.Find("digital-signatures")!.Level);
    }

    private static OdfDocument ReopenFlat(OdfDocument document) {
        var stream = new MemoryStream();
        document.SaveFlatXml(stream);
        stream.Position = 0;
        return OdfDocument.OpenFlatXml(stream);
    }
}
