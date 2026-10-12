using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    [Theory]
    [InlineData("pages-created-edited-v15.4.pages", "pages-pdf", 2, "8")]
    [InlineData("pages-created-headers-v15.4.pages", "pages-pdf", 2, "4")]
    [InlineData("pages-created-headers-edited-v15.4.pages", "pages-pdf", 2, "4")]
    [InlineData("pages-created-layout-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-preserved-paragraphs-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-preserved-paragraphs-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-created-tables-edited-v15.4.pages", "pages-pdf", 1, "10")]
    [InlineData("pages-created-images-v15.4.pages", "pages-pdf", 1, "12")]
    [InlineData("pages-created-images-edited-v15.4.pages", "pages-pdf", 1, "12")]
    [InlineData("pages-created-numbering-v15.4.pages", "pages-pdf", 1, "10")]
    [InlineData("pages-created-numbering-edited-v15.4.pages", "pages-pdf", 1, "10")]
    [InlineData("pages-created-rich-numbering-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-created-rich-numbering-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-list-continuation-v15.4.pages", "pages-pdf", 1, "2")]
    [InlineData("pages-third-level-created.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-third-level-saved-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-third-level-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-preserved-decimal-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-preserved-decimal-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-nine-level-created.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-nine-level-saved-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-nine-level-preserved.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-nine-level-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-decimal-parentheses-created.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-decimal-parentheses-saved-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-decimal-parentheses-preserved.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-decimal-parentheses-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-decimal-closing-created.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-decimal-closing-saved-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-decimal-closing-preserved.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-decimal-closing-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-mixed-font-created.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-mixed-font-saved-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-mixed-font-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-mixed-size-orders-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-mixed-size-created.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-mixed-size-saved-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("pages-mixed-size-edited-v15.4.pages", "pages-pdf", 1, "8")]
    [InlineData("keynote-created-tables-edited-v15.4.key", "keynote-pdf", 3, "13")]
    [InlineData("keynote-created-tables-pdf-export-v15.4.key", "keynote-pdf", 3, "13")]
    [InlineData("keynote-preserved-resized-v15.4.key", "keynote-pdf", 3, "14")]
    [InlineData("keynote-preserved-resized-edited-v15.4.key", "keynote-pdf", 3, "14")]
    [InlineData("keynote-preserved-paragraphs-v15.4.key", "keynote-pdf", 3, "14")]
    [InlineData("keynote-preserved-paragraphs-edited-v15.4.key", "keynote-pdf", 3, "14")]
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
                if (fixture.StartsWith("pages-created-headers", StringComparison.Ordinal))
                    AssertNativePagesPageText(pdf, fixture.Contains("-edited-"));
                else if (fixture == "pages-created-layout-edited-v15.4.pages") AssertNativePagesLayoutText(pages);
                else if (fixture.StartsWith("pages-preserved-paragraphs", StringComparison.Ordinal))
                    AssertNativePagesLayoutText(pages, true, fixture.Contains("-edited", StringComparison.Ordinal));
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
                else if (fixture.StartsWith("pages-created-rich-numbering", StringComparison.Ordinal))
                    AssertCreatedPagesRichList(pdf, fixture.Contains("edited", StringComparison.Ordinal));
                else if (fixture.StartsWith("pages-mixed-font-", StringComparison.Ordinal))
                    AssertMixedFontList(pdf, fixture.Contains("edited", StringComparison.Ordinal));
                else if (fixture.StartsWith("pages-mixed-size-", StringComparison.Ordinal))
                    AssertMixedSizeList(pdf, fixture.Contains("edited", StringComparison.Ordinal),
                        fixture.Contains("orders", StringComparison.Ordinal));
                else if (fixture.StartsWith("pages-third-level", StringComparison.Ordinal)) {
                    string text = string.Concat(Assert.Single(pages).Where(c => !char.IsWhiteSpace(c)));
                    Assert.Contains("6.Resumedgrandchildsix", text);
                    Assert.Contains("8.Parenteightresetsdescendants1.Childresetone1.Grandchildresetafterrootone", text);
                    Assert.Contains("4.Parentrestartfour1.Childafterrestartone1.Grandchildafterrestartone", text);
                    if (fixture.Contains("edited", StringComparison.Ordinal)) Assert.Contains("2.Nativegrandchildtwo", text);
                    Assert.Contains("Endthird-levelprofile", text);
                }
                else if (fixture.StartsWith("pages-preserved-decimal", StringComparison.Ordinal)) {
                    string text = string.Concat(Assert.Single(pages).Where(c => !char.IsWhiteSpace(c)));
                    Assert.Contains(fixture.Contains("-edited", StringComparison.Ordinal) ? "7.Nativeseven" : "7.Editedseven", text);
                    Assert.Contains("1.Childtwo1.Grandchildsix2.Grandchildtwo", text);
                    Assert.Contains("6.Updatedgrandchildsix8.Parenteightresetsdescendants", text);
                    Assert.Contains("4.Parentrestartfour1.Childafterrestartone1.Grandchildafterrestartone", text);
                    Assert.Contains("Endthird-levelprofile", text);
                }
                else if (fixture.StartsWith("pages-decimal-", StringComparison.Ordinal)) {
                    string text = string.Concat(Assert.Single(pages).Where(c => !char.IsWhiteSpace(c)));
                    bool preserved = fixture.Contains("preserved", StringComparison.Ordinal);
                    bool edited = fixture.Contains("edited", StringComparison.Ordinal);
                    bool parentheses = fixture.Contains("parentheses", StringComparison.Ordinal);
                    string Mark(int number) => parentheses ? "(" + number + ")" : number + ")";
                    Assert.Contains(Mark(7) + (edited ? "Rootnovel" : preserved ? "ROOTSEVEN" : "Rootseven"), text);
                    Assert.Contains(Mark(1) + (edited ? "LevelEIGHT" : preserved ? "LEVELEIGHT" : "Leveleight")
                        + Mark(2) + "Deepcontinuedtwo", text);
                    Assert.Contains(Mark(9) + "Deeprestartnine3.Otherrootthree" + Mark(10) + "Resumeddeepten" + Mark(8) + "Rooteight", text);
                    Assert.Contains(Mark(4) + "Rootrestartfour" + Mark(1) + "Restartlevel1", text);
                    Assert.Contains(Mark(1) + "Restartlevel8" + (edited ? Mark(2) + "Nativedeeptwo" : string.Empty) + "Endnine-levelprofile", text);
                }
                else if (fixture.StartsWith("pages-nine-level", StringComparison.Ordinal)) {
                    string text = string.Concat(Assert.Single(pages).Where(c => !char.IsWhiteSpace(c)));
                    bool preserved = fixture.Contains("preserved", StringComparison.Ordinal);
                    bool edited = fixture.Contains("edited", StringComparison.Ordinal);
                    Assert.Contains(edited ? "7.Rootnovel" : preserved ? "7.ROOTSEVEN" : "7.Rootseven", text);
                    Assert.Contains(edited ? "1.LevelEIGHT2.Deepcontinuedtwo" : preserved
                        ? "1.LEVELEIGHT2.Deepcontinuedtwo" : "1.Leveleight2.Deepcontinuedtwo", text);
                    Assert.Contains("9.Deeprestartnine3.Otherrootthree10.Resumeddeepten8.Rooteight", text);
                    Assert.Contains("4.Rootrestartfour1.Restartlevel1", text);
                    Assert.Contains("1.Restartlevel8" + (edited ? "2.Nativedeeptwo" : string.Empty) + "Endnine-levelprofile", text);
                }
                else AssertNativePagesText(pages);
                if (fixture == "pages-created-layout-edited-v15.4.pages"
                    || fixture.StartsWith("pages-preserved-paragraphs", StringComparison.Ordinal)) {
                    var spans = Assert.Single(pdf.Pages).GetTextSpans();
                    Assert.Contains(spans, span => span.Text == "Rich" && span.IsBold && span.FontSize == 14);
                    Assert.Contains(spans, span => span.Text == "plain" && !span.IsBold && span.FontSize == 12);
                }
                break;
            case "keynote-pdf":
                if (fixture.StartsWith("keynote-preserved-resized", StringComparison.Ordinal))
                    AssertKeynoteImages(pdf, fixture.Contains("edited", StringComparison.Ordinal));
                else if (fixture.StartsWith("keynote-preserved-paragraphs", StringComparison.Ordinal))
                    AssertKeynoteImages(pdf, fixture.Contains("-edited", StringComparison.Ordinal), paragraphEdits: true);
                else AssertNativeKeynoteText(pages);
                break;
        }
    }

    private static void AssertKeynoteImages(PdfReadDocument pdf, bool edited, bool paragraphEdits = false) {
        string[] titles = paragraphEdits
            ? ["Edited PNG — image behind text", "Edited alpha PNG", edited ? "Native checked JPEG" : "Edited JPEG"]
            : ["Native PNG — image behind text", "Native alpha PNG", "Native JPEG"];
        for (int index = 0; index < 3; index++) {
            var page = pdf.Pages[index]; Assert.Contains(titles[index], page.ExtractText());
            var image = Assert.Single(page.GetImagePlacements());
            Assert.Equal(paragraphEdits ? (index == 2 ? 80 : 60) : (index == 2 ? 100 : 80), image.X);
            Assert.Equal(paragraphEdits ? (index == 2 ? 600 : 480) : (index == 1 ? 360 : index == 2 && edited ? 421 : 420), image.Width);
            Assert.Equal(paragraphEdits ? (index == 2 ? 240 : 320) : (index == 1 ? 180 : 168), image.Height);
            double top = paragraphEdits ? (index == 0 ? 110 : 120) : (index == 0 ? 125 : 135);
            Assert.Equal(540 - top - image.Height, image.Y);
        }
        Assert.Contains("Covered source label", pdf.Pages[0].ExtractText());
        Assert.Contains("TEXT IN FRONT", pdf.Pages[0].ExtractText());
        Assert.Contains("100% alpha", pdf.Pages[1].ExtractText());
        Assert.Contains("JPEG / stretched frame", pdf.Pages[2].ExtractText());
        var images = pdf.ExtractImages(); Assert.Equal(3, images.Count);
        Assert.All(images, image => { Assert.Equal(96, image.Width); Assert.Equal(64, image.Height); });
        Assert.Single(images, image => image.HasTransparencyMask && image.TransparencyMaskResolved);
        Assert.Contains(pdf.Pages[1].GetTextSpans(),
            span => span.Text.StartsWith("100%", StringComparison.Ordinal) && span.IsBold && span.FontSize == 28);
    }

    [Theory]
    [InlineData("simple.numbers")]
    [InlineData("numbers-created-edited-v14.5.numbers")]
    [InlineData("numbers-preserved-scalars-v14.5.numbers")]
    [InlineData("numbers-preserved-scalars-edited-v14.5.numbers")]
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

    private static void AssertCreatedPagesRichList(PdfReadDocument pdf, bool edited) {
        var page = Assert.Single(pdf.Pages);
        string text = System.Text.RegularExpressions.Regex.Replace(page.ExtractText(), @"\s+", string.Empty);
        int previous = -1;
        foreach (string expected in new[] { "7.Bold first 😀 plain reset", "1.Child italic, underline, strike",
            "5.Combined decoration reset", "Plain gap", "8.Colored first parent eight", "1.Underlined child reset",
            "9.Plain parent nine", "3.Separate Zażółć café", "10.Resumed parent ten" + (edited ? " native edit" : "") }) {
            string wanted = System.Text.RegularExpressions.Regex.Replace(expected, @"\s+", string.Empty);
            int at = text.IndexOf(wanted, previous + 1, StringComparison.Ordinal);
            Assert.True(at > previous, "Missing or reordered rich-list content: " + expected); previous = at;
        }
        var spans = page.GetTextSpans();
        Assert.Contains(spans, span => span.Text == "Bold" && span.IsBold && !span.IsItalic);
        Assert.Contains(spans, span => span.Text == "Child" && !span.IsBold && span.IsItalic);
        Assert.Contains(spans, span => span.Text == "Combined" && span.IsBold && span.IsItalic);
        Assert.All(spans.Where(span => System.Text.RegularExpressions.Regex.IsMatch(span.Text.Trim(), @"^\d+\.$")),
            span => { Assert.False(span.IsBold); Assert.False(span.IsItalic); });
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

    private static void AssertNativePagesPageText(PdfReadDocument pdf, bool edited) {
        string header = "OfficeIMO — " + (edited ? "Native" : "Pages") + " header";
        string footer = "Zażółć café — " + (edited ? "native " : "") + "footer";
        foreach (var page in pdf.Pages) {
            string text = page.ExtractText();
            Assert.Contains(header, text); Assert.Contains(footer, text);
            var spans = page.GetTextSpans();
            Assert.Contains(spans, span => span.Text.Contains("OfficeIMO") && span.IsBold && span.FontSize == 14);
            var footerEnd = Assert.Single(spans, span => span.Text.Contains("café"));
            var footerSpans = spans.Where(span => Math.Abs(span.Y - footerEnd.Y) < 0.001).ToArray();
            Assert.Equal(footer.Replace(" ", string.Empty), string.Concat(footerSpans.Select(span => span.Text)).Replace(" ", string.Empty));
            Assert.All(footerSpans, span => { Assert.True(span.IsItalic); Assert.Equal(12, span.FontSize); });
        }
        string all = string.Join("\n", pdf.Pages.Select(page => page.ExtractText()));
        int previous = -1;
        for (int index = 1; index <= 72; index++) {
            string expected = $"Body line {index:D2} — header/footer creation";
            int at = all.IndexOf(expected, StringComparison.Ordinal);
            Assert.True(at > previous, "Missing or reordered body paragraph: " + expected); previous = at;
            Assert.Equal(-1, all.IndexOf(expected, at + expected.Length, StringComparison.Ordinal));
        }
    }

    private static void AssertNativePagesLayoutText(IReadOnlyList<string> pages, bool preservedEdit = false, bool nativeEdit = false) {
        string text = string.Join(" ", string.Join("\n", pages).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        int previous = -1;
        foreach (string expected in new[] {
            preservedEdit ? "Edited paragraph layout" : "Native paragraph layout", "Relative spacing 1.5: this paragraph has six points before and nine points after.",
            "Start Left Center Right " + (preservedEdit ? "234.56" : "123.45"), "Double spacing: this paragraph has twelve points before and eighteen points after.",
            "Rich layout plain Zażółć 😀", nativeEdit ? "Native checked paragraph." : (preservedEdit ? "Updated" : "Default") + " reset Native layout edit."
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
