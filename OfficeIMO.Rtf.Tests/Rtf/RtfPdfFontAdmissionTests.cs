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

    private static RtfToPdfOptions Options() => new RtfToPdfOptions {
        PdfOptions = new PdfCore.PdfOptions { DefaultFont = PdfCore.PdfStandardFont.Helvetica },
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
        MaximumSystemFontFamilies = 0
    };
}
