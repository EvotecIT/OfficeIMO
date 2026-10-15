using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Measures imported text in the font's design units before PDF width rounding.</summary>
internal sealed class PdfFontUnitLayoutMetrics {
    private readonly Lazy<OfficeOpenTypeReader?> _reader;

    internal PdfFontUnitLayoutMetrics(byte[] data) =>
        _reader = new Lazy<OfficeOpenTypeReader?>(() => OfficeOpenTypeReader.TryCreate(data));

    internal double MeasureAdvanceWidth1000(PdfGlyphRun run, double horizontalScaling = 100D) {
        // Retain the ordinary checked-sum contract even when the layout metric
        // has greater precision than the encoded PDF advances.
        int encodedAdvance = run.TotalAdvanceWidth1000;
        OfficeOpenTypeReader? reader = _reader.Value;
        if (reader == null) return encodedAdvance * horizontalScaling / 100D;

        double advance = 0D;
        OfficeTextShapingResult? source = run.SourceShapingResult;
        for (int index = 0; index < run.Glyphs.Count; index++) {
            int nominal = reader.AdvanceWidth(run.Glyphs[index].GlyphId);
            int positioned = source == null ? checked(nominal + (run.DesignAdvanceAdjustments?[index] ?? 0))
                : checked((source.Glyphs[index].AdvanceWidth ?? nominal) + source.GetAdvanceAdjustment(index));
            // Horizontal character scaling is resolved in design units. Keep
            // the original metric when no scaling was requested.
            advance += ScaleAdvance(positioned, horizontalScaling);
        }
        return advance * 1000D / reader.UnitsPerEm;
    }

    internal static double ScaleAdvance(int advance, double horizontalScaling) =>
        horizontalScaling == 100D ? advance : Math.Truncate(advance * horizontalScaling / 100D);
}
