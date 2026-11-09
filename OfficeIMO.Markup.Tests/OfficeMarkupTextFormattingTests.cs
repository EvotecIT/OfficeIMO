using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Markup;
using OfficeIMO.Markup.Excel;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.Markup.Word;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests.Markup;

public class OfficeMarkupTextFormattingTests {
    [Fact]
    public void SharedConversionResult_PreservesTypedMarkupFailure() {
        var diagnostic = new OfficeMarkupDiagnostic(
            OfficeMarkupDiagnosticSeverity.Error,
            "Invalid markup contract.");
        var result = new OfficeMarkupConversionResult<object>(new object(), new[] { diagnostic });

        OfficeMarkupConversionException exception = Assert.Throws<OfficeMarkupConversionException>(
            () => result.RequireValue());

        Assert.Same(diagnostic, Assert.Single(exception.Diagnostics));
    }

    [Fact]
    public void SharedConversionReport_PreservesExplicitOmissionWithoutChangingLegacyWarningDefaults() {
        var omitted = new OfficeMarkupDiagnostic(
            OfficeMarkupDiagnosticSeverity.Warning,
            "Source block omitted.",
            node: null,
            lossKind: OfficeConversionLossKind.Omission);
        var legacyWarning = new OfficeMarkupDiagnostic(
            OfficeMarkupDiagnosticSeverity.Warning,
            "Source block approximated.");
        var report = new OfficeMarkupConversionReport(new[] { omitted, legacyWarning });

        Assert.Equal(OfficeConversionLossKind.Omission, omitted.LossKind);
        Assert.Equal(OfficeConversionLossKind.Approximation, legacyWarning.LossKind);
        Assert.Equal(
            new[] { OfficeConversionLossKind.Omission, OfficeConversionLossKind.Approximation },
            report.Diagnostics.Select(static diagnostic => diagnostic.LossKind));
        Assert.True(report.HasLoss);
        Assert.ThrowsAny<OfficeMarkupConversionException>(() => report.ThrowIfLoss());
    }

    [Fact]
    public void SharedConversionReport_FailsClosedWhenAnExplicitErrorClaimsNoLoss() {
        var diagnostic = new OfficeMarkupDiagnostic(
            OfficeMarkupDiagnosticSeverity.Error,
            "The target could not be produced.",
            node: null,
            lossKind: OfficeConversionLossKind.None);
        var report = new OfficeMarkupConversionReport(new[] { diagnostic });

        Assert.False(report.Succeeded);
        Assert.True(report.HasLoss);
        Assert.Equal(OfficeConversionLossKind.Failure, diagnostic.LossKind);
        Assert.Equal(OfficeConversionLossKind.Failure, Assert.Single(report.Diagnostics).LossKind);
        Assert.ThrowsAny<OfficeMarkupConversionException>(() => report.ThrowIfLoss());
    }

    [Fact]
    public void NativeMarkupAdapters_ClassifyOmittedBlocksAsOmissions() {
        var excelSource = new OfficeMarkupDocument(OfficeMarkupProfile.Workbook);
        excelSource.Blocks.Add(new OfficeMarkupCodeBlock("text", "omitted"));
        var excel = excelSource.ToExcel(null, out OfficeIMO.OfficeConversionReport conversionReport1);
        var excelReport = (OfficeMarkupConversionReport)conversionReport1;

        var powerPointSource = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        powerPointSource.Blocks.Add(new OfficeMarkupCodeBlock("text", "omitted"));
        var powerPoint = powerPointSource.ToPowerPoint(new MarkupPowerPointExportOptions { RenderMermaidDiagrams = false }, out OfficeIMO.OfficeConversionReport conversionReport2);
        var powerPointReport = (OfficeMarkupPowerPointConversionReport)conversionReport2;

        var wordSource = new OfficeMarkupDocument(OfficeMarkupProfile.Document);
        wordSource.Blocks.Add(new OfficeMarkupExtensionBlock("custom", new Dictionary<string, string>(), "omitted"));
        var word = wordSource.ToWord(new MarkupWordExportOptions { IncludeUnsupportedBlocksAsText = false }, out OfficeIMO.OfficeConversionReport conversionReport3);
        var wordReport = (OfficeMarkupConversionReport)conversionReport3;

        using (excel)
        using (powerPoint)
        using (word) {
            Assert.Contains(excelReport.Diagnostics,
                diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(powerPointReport.Diagnostics,
                diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(wordReport.Diagnostics,
                diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Fact]
    public void PowerPointConversion_PreservesTypedMarkupFailure() {
        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        source.Blocks.Add(new OfficeMarkupSheetBlock("Invalid for presentations"));

        OfficeConversionException exception = Assert.Throws<OfficeConversionException>(() =>
            source.ToPowerPoint(new MarkupPowerPointExportOptions { RenderMermaidDiagrams = false }));
        OfficeMarkupPowerPointConversionReport report = Assert.IsType<OfficeMarkupPowerPointConversionReport>(exception.Report);
        Assert.Contains(report.NativeDiagnostics, diagnostic =>
            diagnostic.Severity == OfficeMarkupDiagnosticSeverity.Error);
    }

    [Fact]
    public void PowerPointThemePreservesHeadingBodyAndExplicitFonts() {
        var slide = new OfficeMarkupSlideBlock("Theme title");
        slide.Blocks.Add(new OfficeMarkupHeadingBlock(2, "Content heading"));
        var custom = new OfficeMarkupHeadingBlock(2, "Explicit heading");
        custom.Attributes["font-family"] = "Consolas";
        slide.Blocks.Add(custom);
        slide.Blocks.Add(new OfficeMarkupParagraphBlock("Body marker"));
        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        source.Blocks.Add(slide);
        using PowerPointDocument presentation = source.ToPowerPoint(new MarkupPowerPointExportOptions { Theme = OfficeTheme.Boardroom });
        using var saved = new MemoryStream();
        presentation.Save(saved);
        saved.Position = 0;
        using PowerPointDocument reopened = PowerPointDocument.Load(saved);
        var runs = reopened.Slides.SelectMany(item => item.TextBoxes).SelectMany(box => box.Paragraphs).SelectMany(paragraph => paragraph.Runs).ToArray();
        Assert.Equal("Georgia", Assert.Single(runs, run => run.Text == "Content heading").FontFamily);
        Assert.Equal("Consolas", Assert.Single(runs, run => run.Text == "Explicit heading").FontFamily);
        Assert.Equal("Aptos", Assert.Single(runs, run => run.Text == "Body marker").FontFamily);
    }

    [Fact]
    public void StyleResolver_ParsesCompleteTypographyAttributes() {
        var block = new OfficeMarkupTextBoxBlock("Mixed case");
        block.Attributes["font-family"] = "Aptos";
        block.Attributes["font-size"] = "18pt";
        block.Attributes["bold"] = "true";
        block.Attributes["italic"] = "true";
        block.Attributes["underline"] = "wavy";
        block.Attributes["strikethrough"] = "double";
        block.Attributes["baseline"] = "superscript";
        block.Attributes["text-case"] = "toggle";
        block.Attributes["small-caps"] = "true";
        block.Attributes["color"] = "336699";
        block.Attributes["highlight"] = "#FFF2CC";

        var document = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        OfficeMarkupResolvedStyle style = Assert.IsType<OfficeMarkupResolvedStyle>(
            OfficeMarkupStyleResolver.Create(document).Resolve(block));

        Assert.Equal("Aptos", style.FontFamily);
        Assert.Equal(18, style.FontSize);
        Assert.True(style.Bold);
        Assert.True(style.Italic);
        Assert.Equal(OfficeTextDecorationStyle.Wavy, style.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Double, style.StrikethroughStyle);
        Assert.Equal(OfficeTextBaseline.Superscript, style.Baseline);
        Assert.Equal(OfficeTextCase.ToggleCase, style.TextCase);
        Assert.True(style.SmallCaps);
        Assert.Equal("#336699", style.TextColor);
        Assert.Equal("#FFF2CC", style.HighlightColor);
    }

    [Fact]
    public void WordExporter_AppliesRootAndNestedListTypographyInsideAndOutsideSections() {
        static OfficeMarkupListBlock CreateList() {
            var root = new OfficeMarkupListBlock(ordered: false);
            root.Attributes["font-family"] = "Consolas";
            root.Attributes["underline"] = "double";
            var rootItem = new OfficeMarkupListItem("Root item");
            var nested = new OfficeMarkupListBlock(ordered: false);
            nested.Attributes["italic"] = "true";
            nested.Attributes["color"] = "336699";
            nested.Items.Add(new OfficeMarkupListItem("Nested item"));
            rootItem.Blocks.Add(nested);
            root.Items.Add(rootItem);
            return root;
        }

        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Document);
        source.Blocks.Add(CreateList());
        var section = new OfficeMarkupSectionBlock("Section");
        section.Blocks.Add(CreateList());
        source.Blocks.Add(section);

        using WordDocument word = source.ToWord(null, out OfficeIMO.OfficeConversionReport conversionReport5);
        WordParagraphInfo[] paragraphs = word.Inspect().Sections
            .SelectMany(item => item.Elements).OfType<WordParagraphInfo>()
            .Where(item => item.Text.Contains("item", StringComparison.Ordinal))
            .ToArray();

        Assert.Equal(4, paragraphs.Length);
        Assert.All(paragraphs.Where(item => item.Text.Contains("Root", StringComparison.Ordinal)), paragraph => {
            WordRunInfo run = Assert.Single(paragraph.Runs);
            Assert.Equal("Consolas", run.FontFamily);
            Assert.Equal(WordUnderlineStyle.Double, run.UnderlineStyle);
        });
        Assert.All(paragraphs.Where(item => item.Text.Contains("Nested", StringComparison.Ordinal)), paragraph => {
            WordRunInfo run = Assert.Single(paragraph.Runs);
            Assert.True(run.Italic);
            Assert.Equal(OfficeColor.Parse("336699"), run.Color);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WordExporter_AppliesBlockTypographyAndCase(bool asynchronous) {
        var paragraph = new OfficeMarkupParagraphBlock("Mixed case");
        AddCommonStyle(paragraph.Attributes, "double", "superscript", "uppercase");
        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Document);
        source.Blocks.Add(paragraph);

        source.Blocks.Add(new OfficeMarkupExtensionBlock("custom", new Dictionary<string, string>(), "omitted"));
        var options = new MarkupWordExportOptions { IncludeUnsupportedBlocksAsText = false };
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-markup-word-report-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "output.docx");
            string destination = Path.Combine(directory, ".", "output.docx");
            OfficeConversionReport fileReport = asynchronous
                ? await source.SaveAsWordAsync(destination, options)
                : source.SaveAsWord(destination, options);
            Assert.Equal(Path.GetFullPath(path), fileReport.OutputPath);
            OfficeMarkupConversionReport native = Assert.IsType<OfficeMarkupConversionReport>(fileReport);
            Assert.Contains(native.NativeDiagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(native.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
            using WordDocument word = WordDocument.Load(path);
            WordParagraphInfo paragraphSnapshot = Assert.IsType<WordParagraphInfo>(
                Assert.Single(Assert.Single(word.Inspect().Sections).Elements));
            WordRunInfo run = Assert.Single(paragraphSnapshot.Runs);

            Assert.Equal("MIXED CASE", run.Text);
            Assert.True(run.Bold);
            Assert.True(run.Italic);
            Assert.Equal(WordUnderlineStyle.Double, run.UnderlineStyle);
            Assert.True(run.DoubleStrikethrough);
            Assert.Equal("Superscript", run.VerticalTextAlignment);
            Assert.Equal("SmallCaps", run.CapsStyle);
            Assert.Equal(OfficeColor.Parse("336699"), run.Color);
            Assert.Equal(OfficeColor.Parse("FFF2CC"), run.BackgroundColor);

            using var output = new MemoryStream();
            OfficeConversionReport streamReport = asynchronous
                ? await source.SaveAsWordAsync(output, options)
                : source.SaveAsWord(output, options);
            Assert.Null(streamReport.OutputPath);
            Assert.True(output.CanWrite);
            OfficeMarkupConversionReport streamNative = Assert.IsType<OfficeMarkupConversionReport>(streamReport);
            Assert.Equal(native.NativeDiagnostics.Select(diagnostic => (diagnostic.Message, diagnostic.LossKind)),
                streamNative.NativeDiagnostics.Select(diagnostic => (diagnostic.Message, diagnostic.LossKind)));
            output.Position = 0;
            using WordDocument streamed = WordDocument.Load(output);
            Assert.Equal("MIXED CASE", Assert.Single(Assert.Single(streamed.Inspect().Sections).Elements.OfType<WordParagraphInfo>()).Text);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExcelExporter_AppliesNativeAccountingUnderlineScriptsAndCase(bool asynchronous) {
        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Workbook);
        source.Blocks.Add(new OfficeMarkupSheetBlock("Data"));
        var range = new OfficeMarkupRangeBlock("A1") { Sheet = "Data" };
        range.Values.Add(new[] { "Mixed case" });
        source.Blocks.Add(range);
        var formatting = new OfficeMarkupFormattingBlock("A1");
        formatting.Attributes["sheet"] = "Data";
        formatting.Attributes["font"] = "Aptos";
        formatting.Attributes["font-size"] = "17pt";
        formatting.Attributes["bold"] = "true";
        formatting.Attributes["italic"] = "true";
        formatting.Attributes["underline"] = "double-accounting";
        formatting.Attributes["strikethrough"] = "true";
        formatting.Attributes["baseline"] = "subscript";
        formatting.Attributes["text-case"] = "toggle";
        formatting.Attributes["color"] = "336699";
        formatting.Attributes["highlight"] = "FFF2CC";
        source.Blocks.Add(formatting);

        source.Blocks.Add(new OfficeMarkupCodeBlock("text", "omitted"));
        var options = new MarkupExcelExportOptions();
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-markup-excel-report-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "output.xlsx");
            string destination = Path.Combine(directory, ".", "output.xlsx");
            OfficeConversionReport fileReport = asynchronous
                ? await source.SaveAsExcelAsync(destination, options)
                : source.SaveAsExcel(destination, options);
            Assert.Equal(Path.GetFullPath(path), fileReport.OutputPath);
            OfficeMarkupConversionReport native = Assert.IsType<OfficeMarkupConversionReport>(fileReport);
            Assert.Contains(native.NativeDiagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(native.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
            using ExcelDocument workbook = ExcelDocument.Load(path);
            ExcelSheet sheet = Assert.Single(workbook.Sheets);
            ExcelCellStyleInfo style = sheet.Cell(1, 1).GetStyle();

            Assert.Equal("mIXED CASE", sheet.Cell(1, 1).GetValue().Value?.ToString());
            Assert.Equal("Aptos", style.FontName);
            Assert.Equal(17D, style.FontSize?.ToPoints());
            Assert.True(style.Bold);
            Assert.True(style.Italic);
            Assert.Equal(ExcelUnderlineStyle.DoubleAccounting, style.UnderlineStyle);
            Assert.True(style.Strikethrough);
            Assert.Equal(ExcelVerticalTextAlignment.Subscript, style.VerticalTextAlignment);
            Assert.Equal("336699", style.FontColorHex);
            Assert.Equal("FFF2CC", style.FillColorHex);

            using var output = new MemoryStream();
            OfficeConversionReport streamReport = asynchronous
                ? await source.SaveAsExcelAsync(output, options)
                : source.SaveAsExcel(output, options);
            Assert.Null(streamReport.OutputPath);
            Assert.True(output.CanWrite);
            OfficeMarkupConversionReport streamNative = Assert.IsType<OfficeMarkupConversionReport>(streamReport);
            Assert.Equal(native.NativeDiagnostics.Select(diagnostic => (diagnostic.Message, diagnostic.LossKind)),
                streamNative.NativeDiagnostics.Select(diagnostic => (diagnostic.Message, diagnostic.LossKind)));
            output.Position = 0;
            using ExcelDocument streamed = ExcelDocument.Load(output);
            Assert.Equal("mIXED CASE", streamed["Data"].Cell(1, 1).GetValue().Value?.ToString());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void PowerPointExporter_AppliesRunTypographyAndCase() {
        var textBox = new OfficeMarkupTextBoxBlock("Mixed case");
        AddCommonStyle(textBox.Attributes, "wavy", "subscript", "lowercase");
        var slide = new OfficeMarkupSlideBlock();
        slide.Blocks.Add(textBox);
        var source = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        source.Blocks.Add(slide);

        using PowerPointDocument presentation = source.ToPowerPoint(null, out OfficeIMO.OfficeConversionReport conversionReport8);
        PowerPointTextBox output = Assert.Single(Assert.Single(presentation.Slides).Shapes.OfType<PowerPointTextBox>(),
            box => box.Text.Contains("mixed case", StringComparison.Ordinal));
        PowerPointTextRun run = Assert.Single(Assert.Single(output.Paragraphs).Runs);

        Assert.Equal("mixed case", run.Text);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal(PowerPointUnderlineStyle.Wavy, run.UnderlineStyle);
        Assert.Equal(PowerPointStrikeStyle.Double, run.StrikeStyle);
        Assert.True(run.BaselinePercent < 0D);
        Assert.Equal(PowerPointCapitalization.SmallCaps, run.Capitalization);
        Assert.Equal(OfficeColor.Parse("336699"), run.Color);
        Assert.Equal(OfficeColor.Parse("FFF2CC"), run.BackgroundColor);
    }

    private static void AddCommonStyle(IDictionary<string, string> attributes, string underline, string baseline, string textCase) {
        attributes["font"] = "Aptos";
        attributes["font-size"] = "18";
        attributes["bold"] = "true";
        attributes["italic"] = "true";
        attributes["underline"] = underline;
        attributes["strikethrough"] = "double";
        attributes["baseline"] = baseline;
        attributes["text-case"] = textCase;
        attributes["small-caps"] = "true";
        attributes["color"] = "336699";
        attributes["highlight"] = "FFF2CC";
    }
}
