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
        ("read the tables", presentation => string.Join("|", presentation.Slides.SelectMany(slide => slide.Tables).Select(table =>
            table.RowCount + "x" + table.ColumnCount + ":" + string.Join(";", Enumerable.Range(0, table.RowCount)
                .SelectMany(row => Enumerable.Range(0, table.ColumnCount).Select(column => table.GetCell(row, column).Text)))))),
        ("add a table with text", presentation => {
            PowerPointSlide slide = presentation.Slides.Count == 0 ? presentation.AddSlide() : presentation.Slides[0];
            PowerPointTable table = slide.AddTable(3, 4, OfficeLength.Points(20), OfficeLength.Points(100), OfficeLength.Points(503), OfficeLength.Points(121));
            for (int row = 0; row < 3; row++) {
                for (int column = 0; column < 4; column++) table.GetCell(row, column).Text = row == 0 ? "Head " + column : "Cell " + row + "," + column + (column == 3 ? "\nsecond" : string.Empty);
            }
            return table.Name + ":" + table.RowCount + "x" + table.ColumnCount + ":" + table.GetCell(2, 3).Text;
        }),
        ("replace the first table cell's text", presentation => {
            PowerPointTable? table = presentation.Slides.SelectMany(slide => slide.Tables).FirstOrDefault();
            if (table == null) return string.Empty;
            table.GetCell(0, 0).Text = "Replaced";
            return table.GetCell(0, 0).Text;
        }),
        ("read and replace the first slide's notes", presentation => {
            if (presentation.Slides.Count == 0) return string.Empty;
            PowerPointNotes notes = presentation.Slides[0].Notes;
            string before = notes.Text;
            notes.Text = "Speaker notes\nsecond line";
            return before + "|" + notes.Text;
        }),
        ("duplicate a slide and edit the copy", presentation => {
            if (presentation.Slides.Count == 0) return string.Empty;
            PowerPointSlide copy = presentation.DuplicateSlide(presentation.Slides.Count - 1);
            PowerPointTextBox? box = copy.TextBoxes.FirstOrDefault();
            if (box != null) box.Text = "Copied";
            return presentation.Slides.Count + ":" + string.Join("|", copy.TextBoxes.Select(text => text.Text));
        }),
        ("add shapes to the first slide", presentation => {
            if (presentation.Slides.Count == 0) return string.Empty;
            PowerPointSlide slide = presentation.Slides[0];
            PowerPointAutoShape rectangle = slide.AddRectangle(OfficeLength.Points(10), OfficeLength.Points(10), OfficeLength.Points(100), OfficeLength.Points(50));
            PowerPointAutoShape ellipse = slide.AddEllipse(OfficeLength.Points(120), OfficeLength.Points(10), OfficeLength.Points(60), OfficeLength.Points(60), "Badge");
            PowerPointAutoShape line = slide.AddLine(OfficeLength.Points(10), OfficeLength.Points(100), OfficeLength.Points(200), OfficeLength.Points(100));
            return string.Join("|", new[] { rectangle, ellipse, line }.Select(shape => shape.Name + ":" + shape.Id));
        }),
        ("format the text boxes", presentation => {
            var read = new List<string>();
            foreach (PowerPointTextBox box in presentation.Slides.SelectMany(slide => slide.TextBoxes).Take(4)) {
                read.Add(box.Bold + ":" + box.Italic + ":" + box.FontSize?.ToPoints() + ":" + box.FontFamily);
                box.Bold = true;
                box.Italic = false;
                box.FontSize = OfficeLength.Points(18);
                box.FontFamily = "Aptos";
                read.Add(box.Bold + ":" + box.Italic + ":" + box.FontSize?.ToPoints() + ":" + box.FontFamily);
            }
            return string.Join("|", read);
        }),
        ("read the layouts and add a slide by layout", presentation => {
            IReadOnlyList<PowerPointSlideLayoutInfo> layouts = presentation.GetSlideLayouts();
            string read = string.Join("|", layouts.Select(layout => layout.LayoutIndex + ":" + layout.Name + ":" + layout.Type + ":" + layout.RelationshipId));
            if (layouts.Count > 1) presentation.AddSlide(layouts[1].Name);
            return read;
        }),
        ("read and set paragraphs", presentation => {
            PowerPointTextBox? box = presentation.Slides.SelectMany(slide => slide.TextBoxes).FirstOrDefault(candidate => candidate.Paragraphs.Count > 0);
            if (box == null) return string.Empty;
            string before = string.Join("|", box.Paragraphs.Select(paragraph => paragraph.Text));
            box.Paragraphs[0].Text = "First paragraph\nwith a break";
            return before + "=>" + string.Join("|", box.Paragraphs.Select(paragraph => paragraph.Text));
        }),
        ("read and format runs", presentation => {
            PowerPointTextRun? run = presentation.Slides.SelectMany(slide => slide.TextBoxes).SelectMany(box => box.Paragraphs).SelectMany(paragraph => paragraph.Runs).FirstOrDefault();
            if (run == null) return string.Empty;
            string before = run.Text + ":" + run.Bold + ":" + run.Italic + ":" + run.FontSize?.ToPoints();
            run.Bold = !run.Bold;
            run.Italic = true;
            run.FontSize = OfficeLength.Points(21);
            run.Text = "Run text";
            return before + "=>" + run.Text + ":" + run.Bold + ":" + run.Italic + ":" + run.FontSize?.ToPoints();
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

        PowerPointTextRun run = box.Paragraphs[0].Runs[0];
        run.Bold = true;
        Assert.False(presentation.IsSdkLoaded);

        _ = presentation.OpenXmlDocument;

        Assert.True(presentation.IsSdkLoaded);
        Assert.True(run.Bold);
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
        PowerPointSlide first = presentation.AddSlide();
        first.AddTextBox("First slide");
        first.Notes.Text = "First notes";
        PowerPointSlide second = presentation.AddSlide();
        PowerPointTextBox box = second.AddTextBox("Second slide");
        box.Paragraphs[0].Runs[0].Bold = true;
        box.Name = "Second title";
        presentation.AddSlide().AddTextBox("Third slide");
        return presentation.ToBytes();
    }
}
