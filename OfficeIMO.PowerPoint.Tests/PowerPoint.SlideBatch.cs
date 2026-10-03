using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointSlideBatchTests {
    [Fact]
    public void BatchSlidesRespectCurrentRelationshipsIdsOrderAndSavedPackage() {
        using var presentation = PowerPointPresentation.Create();
        presentation.AddSlide().AddTextBox("First");
        var part = presentation.OpenXmlDocument.PresentationPart!;
        part.Presentation.SlideIdList!.Elements<SlideId>().Single().Id = 1000U;
        int reserved = 1;
        var relationships = part.Parts.Select(pair => pair.RelationshipId).ToHashSet();
        while (relationships.Contains("rId" + reserved)) reserved++;
        string externalId = "rId" + reserved;
        part.AddExternalRelationship("urn:example:external", new Uri("https://example.org/"), externalId);
        var slides = presentation.AddSlides(4);
        for (int index = 0; index < slides.Count; index++) slides[index].AddTextBox("Batch " + index);
        presentation.RemoveSlide(2);
        presentation.AddSlides(1).Single().AddTextBox("Last");
        Assert.Equal(new uint[] { 1000, 1001, 1003, 1004, 1005 },
            part.Presentation.SlideIdList.Elements<SlideId>().Select(slide => slide.Id!.Value));
        Assert.DoesNotContain(part.Presentation.SlideIdList.Elements<SlideId>(),
            slide => slide.RelationshipId?.Value == externalId);
        using var saved = new MemoryStream();
        presentation.Save(saved);
        saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Assert.Equal(new[] { "First", "Batch 0", "Batch 2", "Batch 3", "Last" },
            reopened.Slides.Select(slide => Assert.Single(slide.TextBoxes).Text));
        Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
    }

    [Fact]
    public void InvalidOrCancelledSlideRequestsDoNotCreateOrphanParts() {
        using var presentation = PowerPointPresentation.Create();
        var part = presentation.OpenXmlDocument.PresentationPart!;
        int before = part.Parts.Count();
        Assert.Throws<ArgumentOutOfRangeException>(() => presentation.AddSlide(masterIndex: -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => presentation.AddSlides(2, layoutIndex: int.MaxValue));
        Assert.Throws<ArgumentOutOfRangeException>(() => presentation.AddSlides(-1));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => presentation.AddSlides(2, cancellationToken: cancellation.Token));
        Assert.Equal(before, part.Parts.Count());
        Assert.Empty(presentation.Slides);
        Assert.Empty(presentation.AddSlides(0));
        presentation.AddSlide();
        part.Presentation.SlideIdList!.Elements<SlideId>().Single().Id = 2147483646U;
        int oneSlide = part.Parts.Count();
        Assert.Throws<InvalidOperationException>(() => presentation.AddSlides(2));
        Assert.Equal(oneSlide, part.Parts.Count());
        Assert.Single(presentation.Slides);
        Assert.Single(presentation.AddSlides(1));
        Assert.Throws<InvalidOperationException>(() => presentation.AddSlide());
    }
}
