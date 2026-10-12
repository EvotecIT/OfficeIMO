using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Theory]
    [InlineData("pages-created-edited-v15.4.pages", "pages-pdf", 2, "8")]
    [InlineData("pages-created-layout-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-created-tables-edited-v15.4.pages", "pages-pdf", 1, "10")]
    [InlineData("pages-created-images-v15.4.pages", "pages-pdf", 1, "12")]
    [InlineData("pages-created-images-edited-v15.4.pages", "pages-pdf", 1, "12")]
    [InlineData("pages-created-numbering-v15.4.pages", "pages-pdf", 1, "10")]
    [InlineData("pages-created-numbering-edited-v15.4.pages", "pages-pdf", 1, "10")]
    [InlineData("pages-list-continuation-v15.4.pages", "pages-pdf", 1, "2")]
    [InlineData("keynote-created-tables-edited-v15.4.key", "keynote-pdf", 3, "13")]
    [InlineData("keynote-created-tables-pdf-export-v15.4.key", "keynote-pdf", 3, "13")]
    public async Task Native_content_pdf_routes_retain_every_page_with_caller_fonts(
        string fixture, string route, int expectedPages, string expectedSourceUnits) {
        using var files = new Files(Path.GetExtension(fixture)[1..], "pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", fixture), files.Input, overwrite: true);
        string sourceHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input)));
        OfficeWorkflowRequest request = files.Request(route);
        request.RegisteredConversionSettings = NativePdfSettings();

        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);

        Assert.True(result.Succeeded, result.Summary + "; " + string.Join("; ",
            result.ConversionEvidence?.FidelityDiagnostics.Select(item => item.Code + ": " + item.Message) ?? []));
        OfficeWorkflowConversionEvidence evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
        Assert.Equal("iwork-editable,pdf", evidence.Facts["conversionStages"]);
        Assert.Equal("False", evidence.Facts["partialEditableReconstruction"]);
        Assert.Equal(expectedSourceUnits, evidence.Facts["sourceUnitCount"]);
        Assert.Equal(expectedSourceUnits, evidence.Facts["reconstructedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["omittedSourceUnitCount"]);
        Assert.Equal("0", evidence.Facts["unassessedSourceUnitCount"]);
        Assert.DoesNotContain(evidence.FidelityDiagnostics,
            item => item.LossKind is OfficeConversionLossKind.Omission or OfficeConversionLossKind.Failure);
        Assert.Contains(evidence.FidelityDiagnostics, item => item.Code == "IWORK_RECORD_FIDELITY_UNASSESSED");
        Assert.Contains(evidence.FidelityDiagnostics, item => item.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Contains(result.Diagnostics, item => item.Code == "OutputReopened");
        Assert.Equal(sourceHash, Assert.Single(result.Diagnostics, item => item.Code == "SourceSnapshot").Details["sha256"], ignoreCase: true);
        Assert.Equal(sourceHash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));

        byte[] output = File.ReadAllBytes(files.Output);
        PdfReadDocument pdf = PdfReadDocument.Open(output);
        Assert.Equal(expectedPages, pdf.Pages.Count);
        IReadOnlyList<string> pages = pdf.Pages.Select(page => page.ExtractText()).ToArray();
        Assert.Equal(expectedPages, pages.Count);
        switch (route) {
            case "pages-pdf":
                if (fixture == "pages-created-layout-edited-v15.4.pages") AssertNativePagesLayoutText(pages);
                else if (fixture == "pages-created-tables-edited-v15.4.pages") AssertNativePagesTableText(pages);
                else if (fixture.StartsWith("pages-created-images", StringComparison.Ordinal)) {
                    string text = Assert.Single(pages); int previous = -1;
                    foreach (string expected in new[] { "Before 😀", "Between", "Grid", "42", "After" }) {
                        int at = text.IndexOf(expected, StringComparison.Ordinal); Assert.True(at > previous); previous = at;
                    }
                    var images = pdf.ExtractImages(); Assert.Equal(3, images.Count);
                    Assert.All(images, image => { Assert.Equal(96, image.Width); Assert.Equal(64, image.Height); });
                    Assert.Contains(images, image => image.HasTransparencyMask && image.TransparencyMaskResolved);
                    var placements = Assert.Single(pdf.Pages).GetImagePlacements();
                    Assert.Equal(new double[] { fixture.Contains("edited") ? 144 : 180, 120, 160 }, placements.Select(image => image.Width));
                    Assert.Equal(new double[] { 120, 80, 80 }, placements.Select(image => image.Height));
                }
                else if (fixture == "pages-list-continuation-v15.4.pages") AssertNativePagesListText(pages);
                else if (fixture.StartsWith("pages-created-numbering", StringComparison.Ordinal))
                    AssertCreatedPagesListText(pages, fixture.Contains("edited", StringComparison.Ordinal));
                else AssertNativePagesText(pages);
                break;
            case "keynote-pdf": AssertNativeKeynoteText(pages); break;
        }
    }

    [Theory]
    [InlineData("simple.numbers")]
    [InlineData("numbers-created-edited-v14.5.numbers")]
    [InlineData("numbers-duration-seconds-v14.5.numbers")]
    [InlineData("numbers-duration-milliseconds-v14.5.numbers")]
    public async Task Native_numbers_region_padding_prevents_complete_pdf_publication(string fixture) {
        using var files = new Files("numbers", "pdf");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", fixture), files.Input, overwrite: true);
        string sourceHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input)));
        File.WriteAllText(files.Output, "keep existing");
        var request = files.Request("numbers-pdf");
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        request.RegisteredConversionSettings = NativePdfSettings();

        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);

        Assert.False(result.Succeeded);
        Assert.Equal("keep existing", File.ReadAllText(files.Output));
        Assert.Contains("padding", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
        Assert.Equal(sourceHash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))));
        Assert.Empty(Directory.GetFiles(files.Root, "*.tmp"));
    }

    private static IWorkWorkflowSettings NativePdfSettings() {
        string fontRoot = Path.Combine(AppContext.BaseDirectory, "Fonts");
        var fonts = new OfficeFontFaceCollection();
        foreach (var (name, style) in new[] {
            ("Regular", OfficeFontStyle.Regular), ("Bold", OfficeFontStyle.Bold),
            ("Italic", OfficeFontStyle.Italic), ("BoldItalic", OfficeFontStyle.Bold | OfficeFontStyle.Italic)
        }) fonts.Add("Carlito", File.ReadAllBytes(Path.Combine(fontRoot, "Carlito-" + name + ".ttf")), style);
        fonts.Add("Noto Emoji", File.ReadAllBytes(Path.Combine(fontRoot, "NotoEmoji-VariableFont_wght.ttf")))
            .AddFallbackFamily("Noto Emoji");
        fonts.Add("Noto JP Common", File.ReadAllBytes(Path.Combine(fontRoot, "NotoSansJP-OfficeIMO-Common.ttf")))
            .AddFallbackFamily("Noto JP Common");
        fonts.AddFallbackFamily("Carlito");
        var profile = new OfficeRenderingProfile("native-iwork-pdf", fonts);
        var settings = new IWorkWorkflowSettings();
        settings.PagesPdfOptions.UseRenderingProfile(profile);
        settings.NumbersPdfOptions.UseRenderingProfile(profile);
        settings.KeynotePdfOptions.UseRenderingProfile(profile);
        settings.PagesPdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
        settings.NumbersPdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
        settings.KeynotePdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
        return settings;
    }

    private static void AssertNativePagesText(IReadOnlyList<string> pages) {
        Assert.Contains("OfficeIMO native body document", pages[0]);
        Assert.Contains("Bold 😀 plain edited — Zażółć café", pages[0]);
        string text = string.Join("\n", pages);
        int previous = -1;
        for (int number = 1; number <= 80; number++) {
            string paragraph = $"Paragraph {number:00}: the body flows over pages with editable text.";
            int at = text.IndexOf(paragraph, StringComparison.Ordinal);
            Assert.True(at > previous, "Missing or reordered paragraph: " + paragraph);
            Assert.Equal(-1, text.IndexOf(paragraph, at + paragraph.Length, StringComparison.Ordinal));
            previous = at;
        }
        Assert.Contains("Paragraph 80:", pages[^1]);
    }

    private static void AssertNativePagesLayoutText(IReadOnlyList<string> pages) {
        string text = string.Join(" ", string.Join("\n", pages).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        int previous = -1;
        foreach (string expected in new[] {
            "Native paragraph layout", "Relative spacing 1.5: this paragraph has six points before and nine points after.",
            "Start Left Center Right 123.45", "Double spacing: this paragraph has twelve points before and eighteen points after.",
            "Rich layout plain Zażółć 😀", "Default reset Native layout edit."
        }) {
            int at = text.IndexOf(expected, StringComparison.Ordinal);
            Assert.True(at > previous, "Missing or reordered content: " + expected);
            previous = at;
        }
    }

    private static void AssertNativePagesTableText(IReadOnlyList<string> pages) {
        string text = Assert.Single(pages);
        Assert.True(text.IndexOf("Before 😀", StringComparison.Ordinal) < text.IndexOf("After", StringComparison.Ordinal));
        foreach (string expected in new[] { "Before 😀", "A😀B", "Zażółć", "=SUM(A1:A2)", "20.75", "TRUE", "FALSE", "After" })
            Assert.Contains(expected, text);
    }

    private static void AssertNativePagesListText(IReadOnlyList<string> pages) {
        string text = string.Join("\n", pages);
        int previous = -1;
        foreach (string expected in new[] { "Native decimal list continuation", "7. Initial seven", "8. Initial eight",
            "Non-list interruption", "9. Continued nine", "1. Nested first", "2. Nested second", "10.Parent ten",
            "1. Nested after parent", "5. Nested restart five", "6. Nested restart six", "Nested non-list interruption",
            "7. Nested continued seven", "11.Original parent eleven", "3. Parent restart three", "4. Parent restart four" }) {
            int at = text.IndexOf(expected, StringComparison.Ordinal);
            Assert.True(at > previous, "Missing or reordered numbered content: " + expected);
            previous = at;
        }
    }

    private static void AssertCreatedPagesListText(IReadOnlyList<string> pages, bool edited) {
        string text = System.Text.RegularExpressions.Regex.Replace(Assert.Single(pages), @"\s+", string.Empty);
        var expected = new List<(int Number, string Text)> {
            (7, "Parent seven"), (edited ? 9 : 8, "Parent eight"), (edited ? 10 : 9, "Continued nine"),
            (1, "Child one"), (2, "Child two"), (edited ? 11 : 10, "Parent ten"), (1, "Reset child one"),
            (5, "Child restart five"), (3, "Separate parent three"), (6, "Resumed child six"),
            (11, "Resumed parent eleven"), (1, "Reset child one again"), (4, "Parent restart four"), (5, "Parent five")
        };
        if (edited) expected.Insert(1, (8, "Inserted parent eight"));
        int previous = -1;
        foreach (var item in expected) {
            string wanted = System.Text.RegularExpressions.Regex.Replace(item.Number
                .ToString(System.Globalization.CultureInfo.InvariantCulture) + "." + item.Text, @"\s+", string.Empty);
            int at = text.IndexOf(wanted, previous + 1, StringComparison.Ordinal);
            Assert.True(at > previous, "Missing or reordered numbered content: " + wanted); previous = at;
        }
        Assert.Contains("Betweenlistitems", text);
    }

    private static void AssertNativeKeynoteText(IReadOnlyList<string> pages) {
        foreach (string expected in new[] { "Native editable table", "Product Units Revenue Active", "Widget 42 1200.5 TRUE",
            "Gadget 0 -25.75 FALSE", "=SUM(A1:A2)", "Zażółć 😀", "Edited 😀",
            "Text, numeric and boolean cells; literal equals-sign text" }) Assert.Contains(expected, pages[0]);
        foreach (string expected in new[] { "Multilingual text and mixed stacking order", "A😀B", "Zażółć — café", "0.125", "AFTER TABLE" })
            Assert.Contains(expected, pages[1]);
        Assert.Equal(2, pages[1].Split("REPEATED", StringSplitOptions.None).Length - 1);
        Assert.Contains("Empty editable table", pages[2]);
        Assert.Contains("Created in empty table", pages[2]);
    }
}
