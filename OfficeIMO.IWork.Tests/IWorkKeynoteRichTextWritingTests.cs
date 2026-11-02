using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkKeynoteRichTextWritingTests {
    [Fact]
    public void Creation_example_matches_the_package_qualified_by_native_Keynote() {
        byte[] expected = File.ReadAllBytes(Fixture("keynote-created-rich-v15.4.key"));
        Assert.Equal(expected, RichExample().SaveBytes());
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(expected),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        converted.Report.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); converted.Value.Save(output); output.Position = 0;
        using var saved = PowerPointDocument.Load(output);
        using var apple = PowerPointDocument.Load(Fixture("keynote-created-rich-v15.4.pptx"));
        Assert.Equal(3, saved.Slides.Count);
        Assert.Empty(saved.ValidateDocument());
        Assert.Empty(saved.Slides[2].TextBoxes);
        for (int slide = 0; slide < 2; slide++) {
            PowerPointTextBox actualBox = Assert.Single(saved.Slides[slide].TextBoxes);
            PowerPointTextBox expectedBox = Assert.Single(apple.Slides[slide].TextBoxes);
            PowerPointTextRun[] actual = actualBox.Paragraphs.SelectMany(p => p.Runs).Where(r => r.Text.Length > 0).ToArray();
            PowerPointTextRun[] oracle = expectedBox.Paragraphs.SelectMany(p => p.Runs).Where(r => r.Text.Length > 0).ToArray();
            Assert.Equal(oracle.Select(r => r.Text), actual.Select(r => r.Text));
            Assert.Equal(oracle.Select(r => r.Bold), actual.Select(r => r.Bold));
            Assert.Equal(oracle.Select(r => r.Italic), actual.Select(r => r.Italic));
            Assert.Equal(oracle.Select(r => r.Underline), actual.Select(r => r.Underline));
            Assert.Equal(oracle.Select(r => r.Strikethrough), actual.Select(r => r.Strikethrough));
            Assert.Equal(oracle.Select(r => r.FontSize), actual.Select(r => r.FontSize));
            // Apple's explicit typeface aliases and theme-backed Arial are separately
            // recorded in the independent range manifest and native font evidence.
            Assert.Equal(expectedBox.Left.ToEmus(), actualBox.Left.ToEmus());
            Assert.Equal(expectedBox.Top.ToEmus(), actualBox.Top.ToEmus());
            Assert.Equal(expectedBox.Width.ToEmus(), actualBox.Width.ToEmus());
            Assert.Equal(expectedBox.Height.ToEmus(), actualBox.Height.ToEmus());
        }
    }

    [Fact]
    public void Apple_edit_and_save_retains_Unicode_ranges_and_plain_formatting_reset() {
        IWorkKeynoteProjection deck = IWorkSourceDocument.Open(Fixture("keynote-created-rich-edited-v15.4.key")).ReadKeynote();
        Assert.Equal(3, deck.Slides.Count);
        IWorkTextContent content = Assert.Single(deck.Slides[1].TextBoxes).Content;
        Assert.Equal("Bold 😀 reset\nItalic\nZażółć — café\n", content.PlainText);
        IWorkTextRun[] runs = content.Paragraphs.SelectMany(p => p.Runs).ToArray();
        IWorkTextStyle decorated = Assert.Single(runs, r => r.Text == "Bold 😀").Style!;
        Assert.True(decorated.Bold); Assert.True(decorated.Italic);
        Assert.True(decorated.Underline); Assert.True(decorated.Strikethrough);
        IWorkTextStyle reset = Assert.Single(runs, r => r.Text == " reset").Style!;
        Assert.False(reset.Bold); Assert.False(reset.Italic);
        Assert.False(reset.Underline); Assert.False(reset.Strikethrough);
        Assert.Equal(24, reset.FontSizePoints);
        Assert.Equal("000000", reset.Color!.RgbHex);
        using var conversion = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(Fixture("keynote-created-rich-edited-v15.4.key"),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        conversion.Report.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); conversion.Value.Save(output); output.Position = 0;
        using var reopened = PowerPointDocument.Load(output);
        Assert.Empty(reopened.ValidateDocument());
        Assert.Contains("reset", reopened.Slides[1].TextBoxes.Single().Text);
    }

    [Fact]
    public void Character_ranges_preserve_Unicode_format_resets_and_paragraphs_in_saved_PowerPoint() {
        var document = IWorkKeynoteDocument.Create();
        document.AddSlide().AddRichText(new[] {
            new IWorkTextRunData("Bold 😀", "Arial", 32, "003366", bold: true, italic: true,
                underline: true, strikethrough: true),
            new IWorkTextRunData(" plain", "Arial", 24),
            new IWorkTextRunData("\r\nItalic", "Times New Roman", 30, "008000", italic: true)
        }, 60, 80, 840, 200);
        byte[] bytes = document.SaveBytes();
        var source = IWorkSourceDocument.Open(new MemoryStream(bytes));
        IWorkTextContent content = Assert.Single(source.ReadKeynote().Slides[0].TextBoxes).Content;
        Assert.Equal("Bold 😀 plain\nItalic\n", content.PlainText);
        Assert.Empty(source.ReadKeynote().SourceDeclarationIssues);
        IWorkTextRun[] sourceRuns = content.Paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
        IWorkTextStyle first = Assert.Single(sourceRuns, run => run.Text == "Bold 😀").Style!;
        Assert.True(first.Bold); Assert.True(first.Italic);
        Assert.True(first.Underline); Assert.True(first.Strikethrough);
        IWorkTextStyle reset = Assert.Single(sourceRuns, run => run.Text == " plain").Style!;
        Assert.False(reset.Bold); Assert.False(reset.Italic);
        Assert.False(reset.Underline); Assert.False(reset.Strikethrough);
        Assert.Equal(24, reset.FontSizePoints);
        Assert.Equal("000000", reset.Color!.RgbHex);
        IWorkTextStyle last = Assert.Single(sourceRuns, run => run.Text == "Italic").Style!;
        Assert.Equal("Times New Roman", last.FontName);
        Assert.Equal(30, last.FontSizePoints);
        Assert.Equal("008000", last.Color!.RgbHex);

        using var conversion = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(bytes),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        conversion.Report.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream();
        conversion.Value.Save(output); output.Position = 0;
        using var reopened = PowerPointDocument.Load(output);
        Assert.Empty(reopened.ValidateDocument());
        PowerPointTextRun[] saved = Assert.Single(reopened.Slides[0].TextBoxes).Paragraphs
            .SelectMany(paragraph => paragraph.Runs).ToArray();
        PowerPointTextRun savedFirst = Assert.Single(saved, run => run.Text == "Bold 😀");
        Assert.True(savedFirst.Bold); Assert.True(savedFirst.Italic);
        Assert.True(savedFirst.Underline); Assert.True(savedFirst.Strikethrough);
        PowerPointTextRun savedReset = Assert.Single(saved, run => run.Text == " plain");
        Assert.False(savedReset.Bold); Assert.False(savedReset.Italic);
        Assert.False(savedReset.Underline); Assert.False(savedReset.Strikethrough);
        Assert.Equal(global::OfficeIMO.OfficeLength.Points(24), savedReset.FontSize);
        Assert.Equal("000000", savedReset.Color);
        PowerPointTextRun savedLast = Assert.Single(saved, run => run.Text == "Italic");
        Assert.True(savedLast.Italic); Assert.False(savedLast.Bold);
        Assert.Equal("Times New Roman", savedLast.FontName);
        Assert.Equal("008000", savedLast.Color);
    }

    [Fact]
    public void Added_runs_are_immutable_snapshots_and_empty_ranges_do_not_create_native_offsets() {
        var document = IWorkKeynoteDocument.Create();
        var input = new List<IWorkTextRunData> {
            new("", "Times New Roman", 80, "FF0000", bold: true), new("visible"), new("")
        };
        IWorkKeynoteTextBox box = document.AddSlide().AddRichText(input, 0, 0, 400, 100);
        input.Clear();
        Assert.Single(box.Runs);
        Assert.Equal("visible", box.Text);
        Assert.Equal("Arial", box.FontName);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkTextRunData>)box.Runs).Clear());
        IWorkTextRun run = Assert.Single(IWorkSourceDocument.Open(new MemoryStream(document.SaveBytes()))
            .ReadKeynote().Slides[0].TextBoxes[0].Content.Paragraphs[0].Runs);
        Assert.Equal("visible", run.Text);
        Assert.Equal("Arial", run.Style!.FontName);

        var blank = IWorkKeynoteDocument.Create();
        IWorkKeynoteTextBox blankBox = blank.AddSlide().AddRichText(new[] {
            new IWorkTextRunData("", "Times New Roman", 20), new IWorkTextRunData("")
        }, 0, 0, 400, 100);
        Assert.Single(blankBox.Runs);
        Assert.Equal("Times New Roman", blankBox.FontName);
        byte[] blankBytes = blank.SaveBytes();
        Assert.Equal("\n", IWorkSourceDocument.Open(new MemoryStream(blankBytes))
            .ReadKeynote().Slides[0].TextBoxes[0].Content.PlainText);
    }

    [Theory]
    [InlineData("bold")]
    [InlineData("italic")]
    [InlineData("underline")]
    [InlineData("strike")]
    public void Native_identity_changes_when_only_a_character_decoration_changes(string decoration) {
        var plain = IWorkKeynoteDocument.Create();
        plain.AddSlide().AddText("text", 0, 0, 400, 100);
        var styled = IWorkKeynoteDocument.Create();
        styled.AddSlide().AddRichText(new[] { new IWorkTextRunData("text", bold: decoration == "bold",
            italic: decoration == "italic", underline: decoration == "underline", strikethrough: decoration == "strike") },
            0, 0, 400, 100);
        byte[] plainBytes = plain.SaveBytes(), styledBytes = styled.SaveBytes();
        Assert.NotEqual(DocumentId(plainBytes), DocumentId(styledBytes));
        Assert.Equal(styledBytes, styled.SaveBytes());
    }

    [Fact]
    public void Character_run_budget_preserves_existing_destinations_before_native_encoding() {
        var document = IWorkKeynoteDocument.Create();
        document.AddSlide().AddRichText(new[] { new IWorkTextRunData("one"), new IWorkTextRunData("two", bold: true) },
            0, 0, 400, 100);
        document.AddSlide().AddText("three", 0, 0, 400, 100);
        var limits = new IWorkKeynoteWriteOptions { MaximumTextRuns = 2 };
        using var stream = new MemoryStream(new byte[] { 9, 8, 7 });
        Assert.Throws<InvalidDataException>(() => document.Save(stream, limits));
        Assert.Equal(new byte[] { 9, 8, 7 }, stream.ToArray());
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-KeynoteRichWriting-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string destination = Path.Combine(folder, "existing.key");
            File.WriteAllBytes(destination, new byte[] { 9, 8, 7 });
            Assert.Throws<InvalidDataException>(() => document.Save(destination, overwrite: true, options: limits));
            Assert.Equal(new byte[] { 9, 8, 7 }, File.ReadAllBytes(destination));
            Assert.Single(Directory.GetFiles(folder));
        } finally { Directory.Delete(folder, recursive: true); }
        limits.MaximumTextRuns = 0;
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SaveBytes(limits));
    }

    [Fact]
    public void Invalid_run_collections_fail_without_adding_a_text_box() {
        var slide = IWorkKeynoteDocument.Create().AddSlide();
        Assert.Throws<ArgumentNullException>(() => slide.AddRichText(null!, 0, 0, 400, 100));
        Assert.Throws<ArgumentException>(() => slide.AddRichText(Array.Empty<IWorkTextRunData>(), 0, 0, 400, 100));
        Assert.Throws<ArgumentException>(() => slide.AddRichText(new IWorkTextRunData[] { null! }, 0, 0, 400, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddRichText(new[] { new IWorkTextRunData("text") }, 900, 0, 400, 100));
        Assert.Empty(slide.TextBoxes);
    }

    private static string DocumentId(byte[] bytes) {
        using var package = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        using var reader = new StreamReader(package.GetEntry("Metadata/DocumentIdentifier")!.Open());
        return reader.ReadToEnd();
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", "native-exports", name);

    private static IWorkKeynoteDocument RichExample() {
        var document = IWorkKeynoteDocument.Create();
        document.AddSlide().AddRichText(new[] {
            new IWorkTextRunData("Plain ", "Helvetica Neue"),
            new IWorkTextRunData("bold", "Helvetica Neue", bold: true),
            new IWorkTextRunData(" ", "Helvetica Neue"),
            new IWorkTextRunData("italic", "Helvetica Neue", italic: true),
            new IWorkTextRunData(" ", "Helvetica Neue"),
            new IWorkTextRunData("underlined", "Helvetica Neue", underline: true),
            new IWorkTextRunData(".", "Helvetica Neue")
        }, 60, 100, 840, 150);
        document.AddSlide("E6F2FF").AddRichText(new[] {
            new IWorkTextRunData("Bold 😀", "Arial", 32, "003366", bold: true, italic: true,
                underline: true, strikethrough: true),
            new IWorkTextRunData(" plain", "Arial", 24),
            new IWorkTextRunData("\r\nItalic", "Times New Roman", 30, "008000", italic: true),
            new IWorkTextRunData("\nZażółć — café", "Arial", 24, "660066")
        }, 60, 80, 840, 250);
        document.AddSlide();
        return document;
    }
}
