using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote.Tests;

public sealed class InkFixtureTests {
    [Fact]
    public void LicensedHandwritingFixturePreservesGeometryAndRecognition() {
        OneNoteSection source = OneNoteSectionReader.Read(FixturePath("handwriting_recognition.one"));
        OfficeInkStroke[] sourceStrokes = Strokes(source);
        OfficeInkStroke[] sourceRecognized = sourceStrokes.Where(stroke => !string.IsNullOrWhiteSpace(stroke.RecognizedText)).ToArray();
        OneNoteTag[] sourceTags = Tags(source);

        Assert.Equal(62, sourceStrokes.Length);
        Assert.Equal(24, sourceRecognized.Length);
        Assert.All(sourceStrokes, stroke => Assert.NotEmpty(stroke.Points));
        Assert.All(sourceRecognized, stroke => Assert.NotEmpty(stroke.RecognitionAlternatives));
        Assert.InRange(sourceStrokes[0].Points[1].X, 1D, 2D);
        Assert.InRange(sourceStrokes[0].Points[1].Y, 30D, 32D);
        byte[] regenerated = OneNoteSectionWriter.Write(source, new OneNoteWriterOptions { ValidateRoundTrip = false });
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(regenerated));
        OfficeInkStroke[] resultStrokes = Strokes(roundTrip);

        Assert.Equal(sourceStrokes.Length, resultStrokes.Length);
        Assert.Equal(sourceRecognized.Select(stroke => stroke.RecognizedText), resultStrokes.Where(stroke => !string.IsNullOrWhiteSpace(stroke.RecognizedText)).Select(stroke => stroke.RecognizedText));
        Assert.Equal(sourceRecognized.Select(stroke => string.Join("\u001F", stroke.RecognitionAlternatives)), resultStrokes.Where(stroke => !string.IsNullOrWhiteSpace(stroke.RecognizedText)).Select(stroke => string.Join("\u001F", stroke.RecognitionAlternatives)));
        Assert.Equal(sourceRecognized.Select(stroke => stroke.LanguageId), resultStrokes.Where(stroke => !string.IsNullOrWhiteSpace(stroke.RecognizedText)).Select(stroke => stroke.LanguageId));
        Assert.Equal(sourceStrokes.Select(stroke => stroke.Points.Count), resultStrokes.Select(stroke => stroke.Points.Count));
        for (int strokeIndex = 0; strokeIndex < sourceStrokes.Length; strokeIndex++) {
            for (int pointIndex = 0; pointIndex < sourceStrokes[strokeIndex].Points.Count; pointIndex++) {
                OfficeInkPoint expected = sourceStrokes[strokeIndex].Points[pointIndex];
                OfficeInkPoint actual = resultStrokes[strokeIndex].Points[pointIndex];
                Assert.InRange(Math.Abs(expected.X - actual.X), 0D, 0.000001D);
                Assert.InRange(Math.Abs(expected.Y - actual.Y), 0D, 0.000001D);
            }
        }
        Assert.Equal(sourceTags.Select(TagIdentity), Tags(roundTrip).Select(TagIdentity));

        OneNotePage inkPage = roundTrip.Pages.OrderByDescending(page => OneNoteElementTraversal.Enumerate(page).OfType<OneNoteInk>().Sum(ink => ink.Strokes.Count)).First();
        OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(inkPage);
        Assert.True(snapshot.Drawing.Height > 1500D);
        Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingShape);
        Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
    }

    private static OfficeInkStroke[] Strokes(OneNoteSection section) => section.Pages
        .SelectMany(page => OneNoteElementTraversal.Enumerate(page))
        .OfType<OneNoteInk>()
        .SelectMany(ink => ink.Strokes)
        .ToArray();

    private static OneNoteTag[] Tags(OneNoteSection section) => section.Pages
        .SelectMany(page => OneNoteElementTraversal.Enumerate(page))
        .SelectMany(element => element.Tags)
        .ToArray();

    private static string TagIdentity(OneNoteTag tag) =>
        $"{tag.ActionItemType}:{tag.IsTask}:{tag.Shape}:{tag.IsCheckable}:{tag.Label}";

    private static string FixturePath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Fixtures", fileName);
}
