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

    [Fact]
    public void RemovingALinkedSlideAfterAnEarlierRemovalLeavesNoDanglingLink() {
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck(3)))) {
            presentation.RemoveSlide(2);
            presentation.Slides[0].TextBoxes.Single().Paragraphs[0].Runs[0].SetHyperlink(presentation.Slides[1]);
            presentation.RemoveSlide(1);
            saved = presentation.ToBytes();
        }

        using PresentationDocument document = PresentationDocument.Open(new MemoryStream(saved), false);
        SlidePart remaining = document.PresentationPart!.SlideParts.Single();
        foreach (A.HyperlinkOnClick link in remaining.Slide!.Descendants<A.HyperlinkOnClick>()) {
            Assert.True(remaining.Parts.Any(pair => pair.RelationshipId == link.Id)
                || remaining.HyperlinkRelationships.Any(relationship => relationship.Id == link.Id), "A hyperlink points at a removed relationship.");
        }
        Assert.DoesNotContain(remaining.Parts, pair => pair.OpenXmlPart is SlidePart);
    }

    [Fact]
    public void AnUnparsedHyperlinkTargetKeepsItsTextWhenAssignedAgain() {
        byte[] linked;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck(1, paragraph => paragraph.AddText("Link"))))) {
            presentation.Slides[0].TextBoxes.Single().Paragraphs[0].Runs[0].SetHyperlink("https://example.com/");
            linked = RewriteHyperlinkTarget(presentation.ToBytes(), "https://example.com/", "mailto:one@");
        }

        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(linked))) {
            PowerPointTextRun run = presentation.Slides[0].TextBoxes.Single().Paragraphs[0].Runs[0];
            run.SetHyperlink(run.Hyperlink!);
            Assert.False(presentation.IsSdkLoaded);
            saved = presentation.ToBytes();
        }

        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(saved));
        using var reader = new StreamReader(archive.GetEntry("ppt/slides/_rels/slide1.xml.rels")!.Open());
        string relationships = reader.ReadToEnd();
        Assert.Contains("Target=\"mailto:one@\"", relationships);
        Assert.DoesNotContain("rewritten://", relationships);
    }

    private static byte[] RewriteHyperlinkTarget(byte[] package, string from, string to) {
        var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        using (var archive = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true)) {
            System.IO.Compression.ZipArchiveEntry entry = archive.GetEntry("ppt/slides/_rels/slide1.xml.rels")!;
            string text;
            using (var reader = new StreamReader(entry.Open())) text = reader.ReadToEnd();
            Assert.Contains(from, text);
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry("ppt/slides/_rels/slide1.xml.rels").Open());
            writer.Write(text.Replace(from, to));
        }
        return output.ToArray();
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
