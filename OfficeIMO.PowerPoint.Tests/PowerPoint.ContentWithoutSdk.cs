using OfficeIMO.PowerPoint;
using Xunit;
using static OfficeIMO.Tests.PowerPointUnchangedSaveTests;

namespace OfficeIMO.Tests;

/// <summary>
/// Reading and editing slide content in a presentation whose Open XML SDK package is not open works on the package and
/// gives the values and saved package the slide model gives through the SDK.
/// </summary>
public class PowerPointContentWithoutSdkTests {
    public static IEnumerable<object[]> CorpusPresentations() => OpenXmlSlideTextReaderTests.CorpusPresentations();

    // Each scenario edits a presentation and returns what it read, which both ways must agree on.
    private static readonly (string Name, Func<PowerPointDocument, string> Run)[] Scenarios = {
        ("read the shapes", presentation => string.Join("|", presentation.Slides.SelectMany(slide => slide.Shapes)
            .Select(shape => shape.GetType().Name + ":" + shape.Id + ":" + shape.Name + ":" + shape.AltText + ":" + shape.Hidden + ":"
                + shape.Left.ToEmus() + "," + shape.Top.ToEmus() + "," + shape.Width.ToEmus() + "," + shape.Height.ToEmus()))),
        ("read the text boxes", presentation => string.Join("|", presentation.Slides.SelectMany(slide => slide.TextBoxes).Select(box => box.Text))),
        ("add a slide with a text box", presentation => {
            PowerPointTextBox box = presentation.AddSlide().AddTextBox("Hello\nworld");
            return box.Name + ":" + box.Id + ":" + box.Text;
        }),
        ("add a text box to the first slide", presentation => presentation.Slides.Count == 0 ? string.Empty
            : presentation.Slides[0].AddTextBox("Added", OfficeLength.Points(10), OfficeLength.Points(20), OfficeLength.Points(300), OfficeLength.Points(40)).Name ?? string.Empty),
        ("add images to the first slide", presentation => {
            if (presentation.Slides.Count == 0) return string.Empty;
            PowerPointSlide slide = presentation.Slides[0];
            PowerPointImage wide = slide.AddImage(new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(40, 10)), OfficeIMO.Drawing.OfficeImageFormat.Png,
                OfficeLength.Points(10), OfficeLength.Points(10), OfficeLength.Points(100), OfficeLength.Points(100));
            PowerPointImage square = slide.AddImage(new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(10, 10)), OfficeIMO.Drawing.OfficeImageFormat.Png);
            PowerPointImage again = slide.AddImage(new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(40, 10)), OfficeIMO.Drawing.OfficeImageFormat.Png);
            return string.Join("|", new[] { wide, square, again }.Select(image => image.Name + ":" + image.Left.ToEmus() + "," + image.Top.ToEmus() + "," + image.Width.ToEmus() + "," + image.Height.ToEmus()));
        }),
        ("add a slide with a text box and an image", presentation => {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Caption");
            return slide.AddImage(new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(10, 30)), OfficeIMO.Drawing.OfficeImageFormat.Png).Name ?? string.Empty;
        }),
        ("replace the first text box's text", presentation => {
            PowerPointTextBox? box = presentation.Slides.SelectMany(slide => slide.TextBoxes).FirstOrDefault(candidate => candidate.Text.Length > 0);
            if (box == null) return string.Empty;
            box.Text = "Changed\r\nSecond line\nThird\vline";
            return box.Text;
        })
    };

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void EditsCorpusPresentationsAsTheModelDoes(string relativePath) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));
        Assert.Empty(Differences(source));
    }

    [Fact]
    public void ReadsAndEditsTextBoxesWithoutOpeningTheSdk() {
        byte[] source = CreateDeck();
        foreach ((string name, Func<PowerPointDocument, string> run) in Scenarios) {
            using var presentation = PowerPointDocument.Load(new MemoryStream(source));
            run(presentation);
            presentation.ToBytes();
            Assert.False(presentation.IsSdkLoaded, name);
            Assert.Empty(presentation.Diagnostics);
        }
        Assert.Empty(Differences(source));
    }

    [Fact]
    public void AShapeTheSdkReadsLaterKeepsWhatWasEditedWithoutIt() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()));
        PowerPointTextBox box = presentation.Slides[1].TextBoxes.Single();
        box.Text = "Edited";
        PowerPointTextBox added = presentation.Slides[1].AddTextBox("Added");
        Assert.False(presentation.IsSdkLoaded);

        box.Paragraphs[0].Runs[0].Bold = true;

        Assert.True(presentation.IsSdkLoaded);
        Assert.Equal("Edited", box.Text);
        Assert.Equal("Added", added.Text);
        using var reopened = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes()));
        Assert.Equal(new[] { "Edited", "Added" }, reopened.Slides[1].TextBoxes.Select(text => text.Text));
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

    private static List<string> Differences(byte[] source) {
        var differences = new List<string>();
        foreach ((string name, Func<PowerPointDocument, string> run) in Scenarios) {
            string? modelValue = null, deferredValue = null;
            byte[]? model = null, deferred = null;
            Exception? modelFailure = Record.Exception(() => {
                using var opened = PowerPointDocument.Load(new MemoryStream(source));
                _ = opened.OpenXmlDocument;
                modelValue = run(opened);
                model = opened.ToBytes();
            });
            Exception? deferredFailure = Record.Exception(() => {
                using var loaded = PowerPointDocument.Load(new MemoryStream(source));
                deferredValue = run(loaded);
                deferred = loaded.ToBytes();
            });
            if (modelFailure != null || deferredFailure != null) {
                if (modelFailure?.GetType() != deferredFailure?.GetType()) differences.Add(name + ": model " + modelFailure?.GetType().Name + ", without the SDK " + deferredFailure?.GetType().Name + " " + deferredFailure?.Message);
                continue;
            }
            if (modelValue != deferredValue) differences.Add(name + ": model read '" + modelValue + "', without the SDK '" + deferredValue + "'");
            Exception? packageDifference = Record.Exception(() => AssertSamePackage(source, model!, deferred!, keptAsLoaded: false));
            if (packageDifference != null) differences.Add(name + ": " + packageDifference.Message);
        }
        return differences;
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.AddSlide().AddTextBox("First slide");
        PowerPointSlide second = presentation.AddSlide();
        PowerPointTextBox box = second.AddTextBox("Second slide");
        box.Paragraphs[0].Runs[0].Bold = true;
        box.Name = "Second title";
        presentation.AddSlide().AddTextBox("Third slide");
        return presentation.ToBytes();
    }
}
