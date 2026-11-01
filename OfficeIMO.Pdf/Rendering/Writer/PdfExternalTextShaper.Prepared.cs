using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfExternalTextShaper {
    // Glyph ids are meaningful only within the exact static TrueType program selected by PDF.
    internal static bool MatchesPreparedFont(OfficePreparedTextShaping prepared, PdfTrueTypeFontProgram font) =>
        prepared.Program is OfficeTrueTypeFont source && source.CollectionIndex == null
        && source.UnitsPerEm == font.UnitsPerEm
        && source.FontDataForShaping.SequenceEqual(font.FontDataForInspection);

    internal static PdfTextShowCommand BuildPreparedTextCommand(OfficePreparedTextShaping prepared,
        PdfTrueTypeFontProgram font, double size) {
        prepared.CancellationToken.ThrowIfCancellationRequested();
        if (!MatchesPreparedFont(prepared, font))
            throw new InvalidOperationException("Prepared drawing glyphs do not belong to the selected embedded PDF font.");
        return BuildGlyphRun(prepared.Text, prepared.Result, font.GlyphCount, font.UnitsPerEm,
            font.GetGlyphWidth1000, font.RecordGlyphUsage, includeActualText: false)
            .ToTextShowCommand().WithFontUnitLayoutAdvance(prepared.Measure(size) * 1000D / size);
    }
}
