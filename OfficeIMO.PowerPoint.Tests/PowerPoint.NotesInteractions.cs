using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointNotesInteractionTests {
    [Fact]
    public void NoteRunHyperlinkCleanupPreservesTheOwningSlideBacklink() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        try {
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointSlide target = presentation.AddSlide();
                PowerPointTextRun run = slide.Notes.SetParagraphs(new[] { "Linked note" })
                    .Single().Runs.Single();
                NotesSlidePart notesPart = slide.SlidePart.NotesSlidePart!;
                string backlinkId = notesPart.GetIdOfPart(slide.SlidePart);

                run.SetHyperlink(target);
                string internalRelationshipId = run.Run.RunProperties!
                    .GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>()!.Id!.Value!;
                HyperlinkRelationship internalRelationship = Assert.Single(notesPart.HyperlinkRelationships);
                Assert.Equal(internalRelationshipId, internalRelationship.Id);
                Assert.False(internalRelationship.IsExternal);
                Assert.Equal("#slide-2", run.Hyperlink!.ToString());
                Assert.Equal(backlinkId, notesPart.GetIdOfPart(slide.SlidePart));

                run.SetHyperlink("https://example.test/note");
                Assert.Same(slide.SlidePart, notesPart.SlidePart);
                Assert.Equal(backlinkId, notesPart.GetIdOfPart(slide.SlidePart));

                run.ClearHyperlink();
                Assert.Same(slide.SlidePart, notesPart.SlidePart);
                Assert.Equal(backlinkId, notesPart.GetIdOfPart(slide.SlidePart));
                Assert.Empty(presentation.ValidateDocument());
                presentation.Save();
            }

            using PowerPointPresentation reopened = PowerPointPresentation.Load(path);
            PowerPointSlide actual = reopened.Slides[0];
            Assert.Same(actual.SlidePart, actual.SlidePart.NotesSlidePart!.SlidePart);
            Assert.Equal("Linked note", actual.Notes.Paragraphs.Single().Text);
            Assert.Empty(reopened.ValidateDocument());
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void NoteRunSoundsUseNotesPartRelationshipsAndRoundTrip() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        byte[] clickBytes = CreateWave(1);
        byte[] mouseOverBytes = CreateWave(2);
        try {
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextRun run = slide.Notes.SetParagraphs(new[] { "Sounded note" })
                    .Single().Runs.Single();
                using (var click = new MemoryStream(clickBytes, writable: false)) {
                    run.SetClickSound(click, "Click note sound");
                }
                using (var mouseOver = new MemoryStream(mouseOverBytes, writable: false)) {
                    run.SetMouseOverSound(mouseOver, "Mouse-over note sound");
                }

                NotesSlidePart notesPart = slide.SlidePart.NotesSlidePart!;
                Assert.Equal(clickBytes, run.GetClickSoundBytes());
                Assert.Equal(mouseOverBytes, run.GetMouseOverSoundBytes());
                Assert.Equal(2, notesPart.DataPartReferenceRelationships
                    .OfType<AudioReferenceRelationship>().Count());
                Assert.Empty(slide.SlidePart.DataPartReferenceRelationships
                    .OfType<AudioReferenceRelationship>());
                presentation.Save();
            }

            using (PowerPointPresentation reopened = PowerPointPresentation.Load(path)) {
                PowerPointSlide slide = reopened.Slides.Single();
                PowerPointParagraph paragraph = slide.Notes.Paragraphs.Single();
                PowerPointTextRun run = paragraph.Runs.Single();
                Assert.Equal(clickBytes, run.GetClickSoundBytes());
                Assert.Equal(mouseOverBytes, run.GetMouseOverSoundBytes());

                paragraph.Text = "Replacement";

                Assert.Empty(slide.SlidePart.NotesSlidePart!
                    .DataPartReferenceRelationships.OfType<AudioReferenceRelationship>());
                Assert.Empty(reopened.OpenXmlDocument.DataParts);
                reopened.Save();
            }

            using PresentationDocument package = PresentationDocument.Open(path, false);
            NotesSlidePart savedNotes = package.PresentationPart!.SlideParts.Single()
                .NotesSlidePart!;
            Assert.Empty(savedNotes.DataPartReferenceRelationships
                .OfType<AudioReferenceRelationship>());
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static byte[] CreateWave(byte marker) => new byte[] {
        0x52, 0x49, 0x46, 0x46, 0x24, 0x00, 0x00, 0x00,
        0x57, 0x41, 0x56, 0x45, 0x66, 0x6D, 0x74, 0x20,
        0x10, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00,
        0x40, 0x1F, 0x00, 0x00, 0x40, 0x1F, 0x00, 0x00,
        0x01, 0x00, 0x08, 0x00, 0x64, 0x61, 0x74, 0x61,
        0x01, 0x00, 0x00, 0x00, marker
    };
}
