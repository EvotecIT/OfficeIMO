using OfficeIMO.PowerPoint;
using Xunit;
using static OfficeIMO.Tests.PowerPointUnchangedSaveTests;

namespace OfficeIMO.Tests;

/// <summary>
/// Removing and moving slides in a presentation whose Open XML SDK package is not open yet edits the slide list without
/// opening it, and saves the package the slide model saves.
/// </summary>
public class PowerPointSlideOrderWithoutSdkTests {
    public static IEnumerable<object[]> CorpusPresentations() => OpenXmlSlideTextReaderTests.CorpusPresentations();

    private static readonly (string Name, int MinimumSlides, Action<PowerPointDocument, int> Edit)[] Edits = {
        ("remove the first slide", 1, (presentation, count) => presentation.RemoveSlide(0)),
        ("remove the last slide", 2, (presentation, count) => presentation.RemoveSlide(count - 1)),
        ("remove a middle slide", 3, (presentation, count) => presentation.RemoveSlide(count / 2)),
        ("move the first slide last", 2, (presentation, count) => presentation.MoveSlide(0, count - 1)),
        ("move the last slide first", 2, (presentation, count) => presentation.MoveSlide(count - 1, 0)),
        ("remove the first slide, then move the last one first", 3, (presentation, count) => {
            presentation.RemoveSlide(0);
            presentation.MoveSlide(count - 2, 0);
        })
    };

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void EditsCorpusPresentationsAsTheModelDoes(string relativePath) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));
        Assert.Empty(Differences(source));
    }

    [Fact]
    public void RemovesAndMovesSlidesWithoutOpeningTheSdk() {
        byte[] source = CreateDeck();
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));

        presentation.RemoveSlide(1);
        presentation.MoveSlide(2, 0);

        Assert.Equal(3, presentation.Slides.Count);
        string text = string.Concat(presentation.ExtractMarkdownChunks().Select(chunk => chunk.Text));
        byte[] saved = presentation.ToBytes();
        Assert.False(presentation.IsSdkLoaded);
        Assert.True(text.IndexOf("Fourth", StringComparison.Ordinal) < text.IndexOf("First", StringComparison.Ordinal));
        Assert.DoesNotContain("Second", text);
        Assert.DoesNotContain(ReadParts(saved).Keys, part => part.Contains("notesSlide", StringComparison.Ordinal));
        Assert.Empty(Differences(source));
    }

    [Fact]
    public void KeepsSectionsInLineWithTheSlides() {
        byte[] source;
        using (var created = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            created.AddSection("Opening", 0);
            created.AddSection("Middle", 1);
            created.AddSection("Closing", 3);
            source = created.ToBytes();
        }

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        presentation.MoveSlide(0, 2);
        presentation.RemoveSlide(3);

        Assert.False(presentation.IsSdkLoaded);
        Assert.Equal(new[] { "Middle", "Opening" }, presentation.GetSections().Select(section => section.Name));
        Assert.Empty(Differences(source));
    }

    [Fact]
    public void TheSlidesAfterAnEditAreTheEditedOnes() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()));
        presentation.RemoveSlide(0);
        presentation.MoveSlide(0, 2);

        Assert.Equal(new[] { "Third slide", "Fourth slide", "Second slide" },
            presentation.Slides.Select(slide => slide.TextBoxes.Single().Text));
        Assert.True(presentation.IsSdkLoaded);
    }

    [Theory]
    [InlineData("custom show")]
    [InlineData("comment")]
    [InlineData("link")]
    public void LeavesWhatTheModelUpdatesElsewhereToTheModel(string variant) {
        byte[] source;
        using (var created = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            if (variant == "custom show") created.AddCustomShow("Short", new[] { created.Slides[0], created.Slides[2] });
            if (variant == "comment") created.AddClassicComment(created.Slides[0], new PowerPointCommentAuthor("Reviewer", "RV"), "Check this");
            if (variant == "link") created.Slides[1].TextBoxes.Single().SetHyperlink(created.Slides[0]);
            source = created.ToBytes();
        }

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        presentation.RemoveSlide(0);

        Assert.True(presentation.IsSdkLoaded);
        Assert.Empty(Differences(source, Edits[0]));
    }

    // Opt-in, as the other corpus checks: OFFICEIMO_PPTX_CORPUS names a folder of presentations.
    [Fact]
    public void ExternalCorpusPresentationsEditAsTheModelDoes() {
        string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_PPTX_CORPUS");
        if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) return;
        var failures = new List<string>();
        foreach (string path in Directory.EnumerateFiles(folder, "*.pptx", SearchOption.AllDirectories).OrderBy(file => file, StringComparer.Ordinal)) {
            byte[] source = File.ReadAllBytes(path);
            try {
                using (PowerPointDocument.Load(new MemoryStream(source))) { }
            } catch (Exception) {
                continue;
            }
            failures.AddRange(Differences(source).Select(difference => path + ": " + difference));
        }
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures.Take(40)));
    }

    private static List<string> Differences(byte[] source, params (string Name, int MinimumSlides, Action<PowerPointDocument, int> Edit)[] edits) {
        var differences = new List<string>();
        int count;
        using (var counted = PowerPointDocument.Load(new MemoryStream(source))) count = counted.Slides.Count;
        foreach ((string name, int minimumSlides, Action<PowerPointDocument, int> edit) in edits.Length == 0 ? Edits : edits) {
            if (count < minimumSlides) continue;
            byte[]? model = null, deferred = null;
            Exception? modelFailure = Record.Exception(() => {
                using var opened = PowerPointDocument.Load(new MemoryStream(source));
                _ = opened.OpenXmlDocument;
                edit(opened, count);
                model = opened.ToBytes();
            });
            Exception? deferredFailure = Record.Exception(() => {
                using var loaded = PowerPointDocument.Load(new MemoryStream(source));
                edit(loaded, count);
                deferred = loaded.ToBytes();
            });
            if (modelFailure != null || deferredFailure != null) {
                if (modelFailure?.GetType() != deferredFailure?.GetType()) differences.Add(name + ": model " + modelFailure?.GetType().Name + ", without the SDK " + deferredFailure?.GetType().Name);
                continue;
            }
            Exception? packageDifference = Record.Exception(() => AssertSamePackage(source, model!, deferred!, keptAsLoaded: false));
            if (packageDifference != null) differences.Add(name + ": " + packageDifference.Message);
        }
        return differences;
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.AddSlide().AddTextBox("First slide");
        PowerPointSlide second = presentation.AddSlide();
        second.AddTextBox("Second slide");
        second.Notes.Text = "Speaker note";
        presentation.AddSlide().AddTextBox("Third slide");
        presentation.AddSlide().AddTextBox("Fourth slide");
        return presentation.ToBytes();
    }
}
