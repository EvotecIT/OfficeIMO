using OfficeIMO;
using OfficeIMO.Word;

internal static class PhysicalFontSmoke {
    internal static void Run() {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontSize = OfficeLength.Points(12.5);
        WordParagraph paragraph = document.AddParagraph("Physical font round trip");
        paragraph.FontSize = OfficeLength.Points(10.25);
        paragraph.KerningMinimumFontSize = OfficeLength.Points(8.25);
        using var source = new MemoryStream(document.ToBytes());
        using WordDocument reopened = WordDocument.Load(source);
        WordParagraph loaded = reopened.Paragraphs[0];
        if (loaded.FontSize != OfficeLength.Points(10.5) ||
            loaded.KerningMinimumFontSize != OfficeLength.Points(8.5) ||
            reopened.Settings.FontSize != OfficeLength.Points(12.5)) {
            throw new InvalidOperationException("Physical font sizes lost their native half-point values.");
        }
        loaded.FontSize = null;
        loaded.KerningMinimumFontSize = null;
        using var clearedSource = new MemoryStream(reopened.ToBytes());
        using WordDocument cleared = WordDocument.Load(clearedSource);
        if (cleared.Paragraphs[0].FontSize.HasValue ||
            cleared.Paragraphs[0].KerningMinimumFontSize.HasValue) {
            throw new InvalidOperationException("Clearing physical font overrides did not restore inheritance.");
        }
        Console.WriteLine("PASS | Physical fonts, half-point rounding, settings, and inheritance after reload");
    }
}
