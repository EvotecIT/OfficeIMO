using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Theory]
    [InlineData("pages-created-edited-v15.4.pages", "pages-pdf", 2, "8")]
    [InlineData("numbers-created-edited-v14.5.numbers", "numbers-pdf", 4, "6")]
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
            case "pages-pdf": AssertNativePagesText(pages); break;
            case "numbers-pdf": AssertNativeNumbersText(pages); break;
            case "keynote-pdf": AssertNativeKeynoteText(pages); break;
        }
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

    private static void AssertNativeNumbersText(IReadOnlyList<string> pages) {
        foreach (string expected in new[] { "Item Value Flag", "Apples 20.75 TRUE", "Negative -0.125 FALSE", "Line1\nLine2 12345.6789 TRUE" })
            Assert.Contains(expected, pages[0]);
        foreach (string expected in new[] { "Notes", "Zażółć 😀", "=literal" }) Assert.Contains(expected, pages[1]);
        Assert.Contains("Totals 99", pages[2]);
        Assert.Contains("FALSE", pages[2]);
        Assert.Equal(string.Empty, pages[3].Trim()); // The completely unset 2-by-3 table retains a drawn page.
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
