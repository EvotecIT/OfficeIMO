using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

/// <summary>
/// Runs and paragraphs keep editing the content that is saved when the SDK opens part-way through an edit, follow the
/// SDK's live slide order once it is open, and leave nothing half-done when an edit fails.
/// </summary>
public class PowerPointTextModelBindingTests {
    [Fact]
    public void EveryRunGetsAThemeColorEvenWhenReadingTheThemeOpensTheSdk() {
        byte[] source = CreateDeck(1, paragraph => {
            paragraph.AddText("First");
            paragraph.AddText(" second");
            paragraph.AddText(" third");
        });
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            Assert.False(presentation.IsSdkLoaded);
            // A theme role without a slot reads the theme, which opens the SDK after the first run.
            presentation.Slides[0].TextBoxes.Single().Paragraphs[0].SetColor(OfficeColor.Success);
            saved = presentation.ToBytes();
        }

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        IReadOnlyList<PowerPointTextRun> runs = reopened.Slides[0].TextBoxes.Single().Paragraphs[0].Runs;
        Assert.Equal(3, runs.Count);
        Assert.All(runs, run => Assert.NotNull(run.Color));
    }

    [Fact]
    public void HyperlinksFollowSlideOrderChangedThroughTheSdk() {
        byte[] source = CreateDeck(3);
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        PowerPointTextRun run = presentation.Slides[0].TextBoxes.Single().Paragraphs[0].Runs[0];
        run.SetHyperlink(presentation.Slides[1]);
        PresentationPart part = presentation.OpenXmlDocument.PresentationPart!;
        P.SlideIdList list = part.Presentation!.SlideIdList!;
        P.SlideId second = list.Elements<P.SlideId>().ElementAt(1);

        // Move the linked slide last directly in the SDK.
        second.Remove();
        list.Append(second);

        Assert.Equal("#slide-3", run.Hyperlink!.OriginalString);
        run.SetHyperlink("#slide-2");
        string id = part.SlideParts.First(slide => slide.Slide!.Descendants<A.HyperlinkOnClick>().Any())
            .Slide!.Descendants<A.HyperlinkOnClick>().Single().Id!.Value!;
        SlidePart target = (SlidePart)part.SlideParts.First(slide => slide.Slide!.Descendants<A.HyperlinkOnClick>().Any()).GetPartById(id);
        Assert.Equal(list.Elements<P.SlideId>().ElementAt(1).RelationshipId!.Value, part.GetIdOfPart(target));
    }

    [Fact]
    public void AFieldWhoseConfigurationFailsIsNotLeftInTheParagraph() {
        byte[] source = CreateDeck(1, paragraph => paragraph.AddText("Text"));
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        PowerPointParagraph paragraph = presentation.Slides[0].TextBoxes.Single().Paragraphs[0];

        Assert.Throws<InvalidOperationException>(() => paragraph.AddField("1", "slidenum", null, _ => throw new InvalidOperationException("configure")));

        Assert.DoesNotContain(paragraph.InlineNodes, node => node.Kind == PowerPointParagraphInlineKind.Field);
    }

    [Fact]
    public void ClearingAHyperlinkOfAReadOnlyPresentationSucceeds() {
        byte[] linked;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck(1, paragraph => paragraph.AddText("Link"))))) {
            presentation.Slides[0].TextBoxes.Single().Paragraphs[0].Runs[0].SetHyperlink("https://example.com/");
            linked = presentation.ToBytes();
        }

        using var readOnly = PowerPointDocument.Load(new MemoryStream(linked), new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        PowerPointTextRun run = readOnly.Slides[0].TextBoxes.Single().Paragraphs[0].Runs[0];
        run.ClearHyperlink();

        Assert.Null(run.Hyperlink);
    }

    private static byte[] CreateDeck(int slides, Action<PowerPointParagraph>? fill = null) {
        using var presentation = PowerPointDocument.Create();
        for (int index = 0; index < slides; index++) {
            PowerPointTextBox box = presentation.AddSlide().AddTextBox(string.Empty);
            if (fill != null) {
                PowerPointParagraph paragraph = box.Paragraphs[0];
                paragraph.ClearInlineContent();
                fill(paragraph);
            } else {
                box.Text = "Slide " + (index + 1);
            }
        }
        return presentation.ToBytes();
    }
}
