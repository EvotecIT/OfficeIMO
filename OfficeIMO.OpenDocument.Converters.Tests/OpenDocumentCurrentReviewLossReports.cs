using System;
using System.IO;
using System.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentCurrentReviewLossReportTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public void WordToOdtReportsFlattenedNestedListLevels() {
        using WordDocument source = WordDocument.Create();
        WordList list = source.AddListNumbered();
        list.AddItem("Parent");
        list.AddItem("Child", 1);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocument();
        using OdtDocument target = conversion.Document;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "list-levels" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void OdtToWordReportsHeaderAndFooterImagesAsSkipped() {
        using OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph("Logo").AddImage(TinyPng, "header.png",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        OdfConversionResult<WordDocument> conversion = source.ToWordDocument();
        using WordDocument target = conversion.Document;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "images" &&
            mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Count == 1);
    }

    [Fact]
    public void OdpToPowerPointReportsFlattenedListsAndMixedRuns() {
        using OdpPresentation source = OdpPresentation.Create();
        OdpTextBox textBox = source.AddSlide("Text").AddTextBox(
            OdfRect.FromCentimeters(1, 1, 8, 4), null, "Content");
        OdpParagraph mixed = textBox.AddParagraph("Plain ");
        mixed.AddRun("Bold").Bold = true;
        textBox.AddList().AddItem("Bullet");

        OdfConversionResult<PowerPointPresentation> conversion = source.ToPowerPointPresentation();
        using PowerPointPresentation target = conversion.Document;

        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "text-lists" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "inline-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Contains("Plain Bold", target.Slides.Single().TextBoxes.Single().Text, StringComparison.Ordinal);
    }
}
