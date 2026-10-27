using OfficeIMO.Drawing;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfPdfFontAdmissionTests {
    private const string ExcludedFamily = "Portable Excluded Note Family";

    [Theory]
    [InlineData("detached-excluded")]
    [InlineData("detached-included")]
    [InlineData("run-excluded")]
    [InlineData("generated-excluded")]
    [InlineData("hidden-reference")]
    [InlineData("disabled-table")]
    [InlineData("merged-cell")]
    [InlineData("flattened-nested-table")]
    public void UnpaintedStyledContentDoesNotAdmitItsFont(string scenario) {
        RtfDocument document = RtfDocument.Create();
        int bodyFont = document.AddFont("Helvetica");
        document.Settings.SetDefaultFont(bodyFont);
        int excludedFont = document.AddFont(ExcludedFamily);
        RtfParagraph body = document.AddParagraph("VISIBLE AB");
        var note = new RtfNote(RtfNoteKind.Footnote);
        note.AddParagraph().AddText("NOTEONLY").FontId = excludedFont;
        var options = Options();
        bool markerVisible = scenario is "run-excluded" or "generated-excluded";
        switch (scenario) {
            case "detached-excluded":
            case "detached-included":
                document.AddNote(RtfNoteKind.Footnote).AddParagraph().AddText("NOTEONLY").FontId = excludedFont;
                options.IncludeNotes = scenario == "detached-included";
                break;
            case "run-excluded":
                body.AddText(" [1]").Note = note;
                options.IncludeNotes = false;
                break;
            case "generated-excluded":
                body.AddNoteReference(note, " [1]");
                options.IncludeNotes = false;
                break;
            case "hidden-reference":
                RtfRun hidden = body.AddText("HIDDEN");
                hidden.Hidden = true;
                hidden.FontId = excludedFont;
                hidden.Note = note;
                break;
            case "disabled-table":
                document.AddTable(1, 1).Rows[0].Cells[0].AddParagraph().AddText("HIDDEN").FontId = excludedFont;
                options.IncludeTables = false;
                break;
            case "merged-cell":
                RtfTable table = document.AddTable(1, 2);
                table.Rows[0].Cells[0].HorizontalMerge = RtfTableCellMerge.First;
                table.Rows[0].Cells[0].AddParagraph("CELL");
                table.Rows[0].Cells[1].HorizontalMerge = RtfTableCellMerge.Continue;
                table.Rows[0].Cells[1].AddParagraph().AddText("HIDDEN").FontId = excludedFont;
                break;
            case "flattened-nested-table":
                RtfTableCell outer = document.AddTable(1, 1).Rows[0].Cells[0];
                outer.AddTable(1, 1).Rows[0].Cells[0].AddParagraph().AddText("FLAT").FontId = excludedFont;
                break;
            default: throw new ArgumentOutOfRangeException(nameof(scenario));
        }
        byte[] sourceBefore = document.ToBytes();

        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
        byte[] bytes = result.ToBytes();
        using PdfPigDocument actual = PdfPigDocument.Open(bytes);
        Assert.Equal(1, actual.NumberOfPages);
        string text = string.Concat(actual.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("VISIBLE AB", text);
        Assert.DoesNotContain("NOTEONLY", text);
        Assert.DoesNotContain("HIDDEN", text);
        if (markerVisible) Assert.Contains("[1]", text);
        if (scenario == "flattened-nested-table") Assert.Contains("FLAT", text);
        Assert.All(actual.GetPage(1).Letters, letter => Assert.Contains("Helvetica", letter.FontName));
        Assert.DoesNotContain(result.Warnings, warning => warning.Details.TryGetValue("fontFamily", out string? family)
            && family == ExcludedFamily);
        Assert.Equal(markerVisible, result.Warnings.Any(warning => warning.Code == "NotesSkipped"));
        if (scenario.StartsWith("detached", StringComparison.Ordinal)) result.RequireNoLoss();
        Assert.Equal(sourceBefore, document.ToBytes());
    }

    [Theory]
    [InlineData("run")]
    [InlineData("generated")]
    [InlineData("field")]
    public void IncludedReferencedNoteBodiesRetainRealFontLoss(string reference) {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont("Helvetica"));
        var note = new RtfNote(RtfNoteKind.Footnote);
        note.AddParagraph().AddText("NOTEONLY").FontId = document.AddFont(ExcludedFamily);
        RtfParagraph body = document.AddParagraph("BODY");
        if (reference == "field") body = body.AddField("MERGEFIELD sample").Result;
        if (reference == "generated") body.AddNoteReference(note, " [1]");
        else body.AddText(" [1]").Note = note;
        byte[] sourceBefore = document.ToBytes();

        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(Options());
        using PdfPigDocument actual = PdfPigDocument.Open(result.ToBytes());
        Assert.Contains("NOTEONLY", string.Concat(actual.GetPage(1).Letters.Select(letter => letter.Value)));
        Assert.Single(result.Warnings, warning => warning.Code == "FontFamilySubstituted"
            && warning.Details.TryGetValue("fontFamily", out string? family) && family == ExcludedFamily);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Equal(sourceBefore, document.ToBytes());
    }

    [Theory]
    [InlineData("body")]
    [InlineData("table")]
    [InlineData("header")]
    [InlineData("field-empty")]
    [InlineData("field-visible")]
    [InlineData("styled-tab")]
    public void ExplicitRunsDoNotAdmitUnusedParagraphDefaults(string scenario) {
        RtfDocument document = RtfDocument.Create();
        int visible = document.AddFont("Helvetica");
        document.Settings.SetDefaultFont(visible);
        RtfStyle unused = document.AddStyle(7, "Unused paragraph font");
        unused.FontId = document.AddFont(ExcludedFamily);
        unused.FontSize = 12;
        RtfParagraph paragraph = scenario switch {
            "table" => document.AddTable(1, 1).Rows[0].Cells[0].AddParagraph(),
            "header" => document.AddHeader().AddParagraph(),
            "field-empty" or "field-visible" => document.AddParagraph("BODY").AddField("MERGEFIELD sample").Result,
            _ => document.AddParagraph()
        };
        if (scenario == "header") document.AddParagraph("BODY");
        paragraph.SetStyle(7);
        if (scenario != "field-empty") {
            RtfRun run = paragraph.AddText(scenario == "styled-tab" ? "\t" : "VISIBLE");
            run.FontId = visible;
            run.FontSize = 12;
        }
        byte[] before = document.ToBytes();

        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(Options());
        using PdfPigDocument actual = PdfPigDocument.Open(result.ToBytes());
        Assert.Equal(1, actual.NumberOfPages);
        if (scenario != "styled-tab") Assert.NotEmpty(actual.GetPage(1).Letters);
        Assert.All(actual.GetPage(1).Letters, letter => Assert.Contains("Helvetica", letter.FontName));
        Assert.DoesNotContain(result.Warnings, warning => warning.Details.TryGetValue("fontFamily", out string? family)
            && family == ExcludedFamily);
        result.RequireNoLoss();
        Assert.Equal(before, document.ToBytes());
    }

    [Theory]
    [InlineData("empty")]
    [InlineData("empty-before-page-break")]
    public void RealParagraphStrutsRetainTheirFontLoss(string scenario) {
        RtfDocument document = RtfDocument.Create();
        int visible = document.AddFont("Helvetica");
        document.Settings.SetDefaultFont(visible);
        document.AddStyle(7, "Real paragraph font").FontId = document.AddFont(ExcludedFamily);
        RtfParagraph paragraph = document.AddParagraph().SetStyle(7);
        if (scenario == "empty-before-page-break") {
            paragraph.AddPageBreak();
            paragraph.AddText("BODY").FontId = visible;
        }
        byte[] before = document.ToBytes();

        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(Options());
        using PdfPigDocument actual = PdfPigDocument.Open(result.ToBytes());
        Assert.True(actual.NumberOfPages >= 1);
        Assert.Single(result.Warnings, warning => warning.Code == "FontFamilySubstituted"
            && warning.Details.TryGetValue("fontFamily", out string? family) && family == ExcludedFamily);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Equal(before, document.ToBytes());
    }

    [Theory]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("field")]
    [InlineData("converted")]
    [InlineData("failed-conversion")]
    [InlineData("disabled-image")]
    [InlineData("disabled-story")]
    [InlineData("plain-story")]
    public void RunningImageFontAdmissionMatchesItsActualInlineProjection(string scenario) {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont("Helvetica"));
        document.AddParagraph("BODY");
        RtfStyle unused = document.AddStyle(7, "Image-only paragraph");
        unused.FontId = document.AddFont(ExcludedFamily);
        unused.FontSize = 12;
        RtfParagraph paragraph = (scenario == "footer" ? document.AddFooter() : document.AddHeader())
            .AddParagraph().SetStyle(7);
        if (scenario == "field") paragraph = paragraph.AddField("MERGEFIELD picture").Result.SetStyle(7);
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(12, 8, OfficeColor.Red));
        bool converted = scenario is "converted" or "failed-conversion";
        int imageCount = scenario == "converted" ? 2 : 1;
        for (int index = 0; index < imageCount; index++) {
            RtfImage image = paragraph.AddImage(converted ? RtfImageFormat.Emf : RtfImageFormat.Png,
                converted ? new byte[] { 1, 2, 3 } : png);
            image.DesiredWidthTwips = 240;
            image.DesiredHeightTwips = 160;
        }
        var options = Options();
        options.IncludeImages = scenario != "disabled-image";
        options.IncludeHeaderFooters = scenario != "disabled-story";
        if (scenario == "plain-story") options.PdfOptions!.HeaderAlign = PdfCore.PdfAlign.Left;
        int calls = 0;
        options.ImageConverter = _ => { calls++; return scenario == "failed-conversion" ? null : png; };
        byte[] before = document.ToBytes();

        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        var read = PdfCore.PdfReadDocument.Open(pdf);
        Assert.Single(read.Pages);
        bool painted = options.IncludeImages && options.IncludeHeaderFooters &&
            scenario is not ("plain-story" or "failed-conversion");
        Assert.Equal(painted ? imageCount : 0,
            read.Pages[0].GetImagePlacements().Count);
        Assert.Contains("BODY", read.ExtractText());
        Assert.Equal(converted ? imageCount : 0, calls);
        bool realStrut = scenario is "disabled-image" or "failed-conversion";
        Assert.Equal(realStrut, result.Warnings.Any(warning => warning.Code == "FontFamilySubstituted"
            && warning.Details.TryGetValue("fontFamily", out string? family) && family == ExcludedFamily));
        if (scenario == "field") {
            Assert.Single(result.Warnings, warning => warning.Code == "HeaderFooterFieldResultUsed");
            Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        } else if (painted) {
            result.RequireNoLoss();
        }
        if (scenario == "failed-conversion") Assert.Single(result.Warnings, warning => warning.Code == "ImageConversionFailed");
        if (scenario == "converted") {
            // Replaying a completed PDF does not rerun callbacks, while a new
            // conversion with the same caller options must prepare its own data.
            result.ToBytes();
            Assert.Equal(imageCount, calls);
            document.ToPdfDocumentResult(options).ToBytes();
            Assert.Equal(imageCount * 2, calls);
        }
        Assert.Equal(before, document.ToBytes());
    }

    private static RtfToPdfOptions Options() => new RtfToPdfOptions {
        PdfOptions = new PdfCore.PdfOptions { DefaultFont = PdfCore.PdfStandardFont.Helvetica },
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
        MaximumSystemFontFamilies = 0
    };
}
