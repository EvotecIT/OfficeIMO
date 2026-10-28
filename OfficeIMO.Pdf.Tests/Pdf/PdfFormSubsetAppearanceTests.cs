using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFormSubsetAppearanceTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void FillingBeyondTheOriginalSubsetKeepsTextVisibleAfterFlattening(bool cff, bool appendRevision) {
        string fontPath = cff
            ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!
            : Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf");
        Assert.True(File.Exists(fontPath), "The bundled form font fixture must be available.");
        using PdfDocument source = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false })
            .UseFontFamily("Form body", fontPath)
            .Paragraph(p => p.Text("ABC abc"))
            .TextField("inspector", width: 280)
            .TextField("date", width: 180)
            .TextField("notes", width: 450, height: 48, style: new PdfFormFieldStyle { IsMultiline = true })
            .TextField("code", width: 180, style: new PdfFormFieldStyle { IsComb = true, MaxLength = 4 })
            .ChoiceField("disposition", new[] { "Awaiting", "Return to pool" }, value: "Awaiting");
        byte[] original = source.ToBytes();
        var values = new Dictionary<string, string> {
            ["inspector"] = "Marta Nowak",
            ["date"] = "2026-09-17",
            ["notes"] = "Ready for the next booking.",
            ["code"] = "ZX91",
            ["disposition"] = "Return to pool"
        };
        using PdfDocument filled = appendRevision
            ? PdfDocument.Load(original).Forms.AppendRevision(values,
                new PdfIncrementalFormFieldUpdateOptions { GenerateAppearanceStreams = true })
            : source.Forms.Fill(values);
        Assert.Equal("Marta Nowak", filled.Inspect().FormFields.Single(field => field.Name == "inspector").Value);
        using PdfDocument flattened = filled.Forms.Flatten();
        // Read the painted appearance independently: logical field values alone cannot prove visibility.
        using var independent = UglyToad.PdfPig.PdfDocument.Open(flattened.ToBytes());
        string text = string.Join("\n", independent.GetPages().Select(page => page.Text));
        Assert.Contains("Marta Nowak", text, StringComparison.Ordinal);
        Assert.Contains("2026-09-17", text, StringComparison.Ordinal);
        Assert.Contains("Ready for the next booking.", text, StringComparison.Ordinal);
        Assert.Contains("Return to pool", text, StringComparison.Ordinal);
        // Comb cells are separate text placements; preserve their characters in reading order.
        Assert.Contains("ZX91", string.Concat(text.Where(character => !char.IsWhiteSpace(character))), StringComparison.Ordinal);
        Assert.Equal(original, source.ToBytes());
    }
}
