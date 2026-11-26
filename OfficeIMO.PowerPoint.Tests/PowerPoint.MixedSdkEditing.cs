using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

/// <summary>
/// Callers who mix the library with direct Open XML SDK calls reach the SDK through <see cref="PowerPointDocument.OpenXmlDocument"/>.
/// The first access builds the SDK package from everything edited so far; from then on that package holds the document,
/// so SDK edits and later library edits all reach the saved file.
/// </summary>
public class PowerPointMixedSdkEditingTests {
    [Fact]
    public void LibraryAndSdkEditsOfALoadedPresentationAllReachTheSave() {
        byte[] source = CreateDeck();
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            presentation.Slides[0].TextBoxes.Single().Text = "Edited by OfficeIMO";
            presentation.AddSlide().AddTextBox("Added by OfficeIMO");
            presentation.ReplaceAll("Third", "Last");
            Assert.False(presentation.IsSdkLoaded);

            PresentationDocument document = presentation.OpenXmlDocument;
            SlidePart first = FirstSlidePart(document);
            first.Slide!.Descendants<A.Text>().Single(text => text.Text == "Edited by OfficeIMO").Text = "Edited by the SDK";
            document.PackageProperties.Title = "Title set through the SDK";

            presentation.Slides[1].TextBoxes.Single().Text = "Second, edited again by OfficeIMO";
            presentation.Slides[3].AddTextBox("Added after the SDK opened");
            saved = presentation.ToBytes();
        }

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        Assert.Equal(new[] { "Edited by the SDK", "Second, edited again by OfficeIMO", "Last slide", "Added by OfficeIMO|Added after the SDK opened" },
            reopened.Slides.Select(slide => string.Join("|", slide.TextBoxes.Select(box => box.Text))));
        Assert.Equal("Title set through the SDK", reopened.OpenXmlDocument.PackageProperties.Title);
    }

    [Fact]
    public void ACreatedPresentationKeepsLibraryAndSdkEdits() {
        byte[] saved;
        using (var presentation = PowerPointDocument.Create(new MemoryStream())) {
            PowerPointTextBox box = presentation.AddSlide().AddTextBox("Created");
            box.Bold = true;

            P.Shape shape = FirstSlidePart(presentation.OpenXmlDocument).Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single();
            shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Description = "Described through the SDK";

            box.Text = "Created, then edited";
            presentation.AddSlide().AddTextBox("Second");
            saved = presentation.ToBytes();
        }

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        PowerPointTextBox first = reopened.Slides[0].TextBoxes.Single();
        Assert.Equal("Created, then edited", first.Text);
        Assert.True(first.Bold);
        Assert.Equal("Described through the SDK", first.AltText);
        Assert.Equal("Second", reopened.Slides[1].TextBoxes.Single().Text);
    }

    [Fact]
    public void ObjectsTakenBeforeTheSdkOpensEditTheSdkPackageAfterwards() {
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            PowerPointSlide slide = presentation.Slides[1];
            PowerPointTextBox box = slide.TextBoxes.Single();
            PowerPointTextRun run = box.Paragraphs[0].Runs[0];

            _ = presentation.OpenXmlDocument;
            run.Italic = true;
            box.Name = "Renamed after the SDK opened";
            slide.Hidden = true;
            saved = presentation.ToBytes();
        }

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        PowerPointTextBox reread = reopened.Slides[1].TextBoxes.Single();
        Assert.True(reread.Paragraphs[0].Runs[0].Italic);
        Assert.Equal("Renamed after the SDK opened", reread.Name);
        Assert.True(reopened.Slides[1].Hidden);
    }

    [Fact]
    public void ADocumentThatNeverAsksForTheSdkNeverLoadsIt() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()));
        presentation.Slides[0].TextBoxes.Single().Text = "Edited";
        presentation.AddSlide().AddTextBox("Added");
        presentation.RemoveSlide(1);
        _ = presentation.ToBytes();

        Assert.False(presentation.IsSdkLoaded);
        Assert.Empty(presentation.Diagnostics);
    }

    private static SlidePart FirstSlidePart(PresentationDocument document) {
        PresentationPart part = document.PresentationPart!;
        P.SlideId first = part.Presentation!.SlideIdList!.Elements<P.SlideId>().First();
        return (SlidePart)part.GetPartById(first.RelationshipId!.Value!);
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.AddSlide().AddTextBox("First slide");
        presentation.AddSlide().AddTextBox("Second slide");
        presentation.AddSlide().AddTextBox("Third slide");
        return presentation.ToBytes();
    }
}
