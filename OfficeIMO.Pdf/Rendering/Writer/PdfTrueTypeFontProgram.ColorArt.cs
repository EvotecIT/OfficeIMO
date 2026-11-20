using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    private const int MaximumColorGlyphPoints = 4_000_000;
    private const int MaximumCachedColorGlyphs = 4096;

    // Shared by every document fork of this program: the Core reader of a large emoji font is parsed once.
    private readonly Lazy<ColorGlyphSource> _colorGlyphSource;

    private static Lazy<ColorGlyphSource> CreateColorGlyphSource(byte[] data, IReadOnlyDictionary<int, int> cmap) =>
        new(() => new ColorGlyphSource(data, cmap), LazyThreadSafetyMode.ExecutionAndPublication);

    // COLR layers, PNG glyphs and unreadable color artwork come from the Core reader of the same font bytes.
    private sealed class ColorGlyphSource {
        private readonly Dictionary<int, PdfColorGlyphArt?> _art = new();

        internal ColorGlyphSource(byte[] data, IReadOnlyDictionary<int, int> cmap) {
            if (!HasAnyTable(data, "COLR", "CBDT", "sbix", "SVG ")) return;
            Program = OfficeTrueTypeFont.TryLoad(data);
            if (Program == null) return;
            BlankGlyphs = new HashSet<int>();
            foreach (KeyValuePair<int, int> mapping in cmap) {
                if (mapping.Value > 0 && (OfficeManagedTextShaper.IsNonPaintingControl(mapping.Key)
                    || mapping.Key <= char.MaxValue && char.IsWhiteSpace((char)mapping.Key))) BlankGlyphs.Add(mapping.Value);
            }
        }

        internal OfficeTrueTypeFont? Program { get; }

        /// <summary>Glyphs that whitespace and nonpainting controls map to: blank by design.</summary>
        internal HashSet<int>? BlankGlyphs { get; }

        internal bool HasPaintedArt => Program != null && (Program.HasColorGlyphs || Program.HasBitmapGlyphs);

        internal PdfColorGlyphArt? GetColorArt(int glyphId) {
            lock (_art) {
                if (_art.TryGetValue(glyphId, out PdfColorGlyphArt? cached)) return cached;
            }
            PdfColorGlyphArt? art = null;
            OfficeTrueTypeFont program = Program!;
            if (program.TryGetGlyphColorLayers(glyphId, OfficeColor.Black, MaximumColorGlyphPoints, default, out var layers)) {
                art = new PdfColorGlyphArt(program.UnitsPerEm, layers.Select(layer =>
                    new PdfColorGlyphLayer(PdfColorGlyphArt.ToPathOperators(layer.Contours), layer.Color, layer.Foreground)).ToArray());
            } else if (program.TryGetBitmapGlyph(glyphId, out OfficeBitmapGlyphImage? bitmap)) {
                art = new PdfColorGlyphArt(program.UnitsPerEm, Array.Empty<PdfColorGlyphLayer>(), bitmap);
            }
            lock (_art) {
                if (_art.Count >= MaximumCachedColorGlyphs) _art.Clear();
                _art[glyphId] = art;
            }
            return art;
        }

        internal PdfColorGlyphArt GetOutlineArt(int glyphId) {
            OfficeTrueTypeFont program = Program!;
            string path = program.TryGetGlyphOutline(glyphId, MaximumColorGlyphPoints, default, out List<List<OfficePoint>> contours)
                ? PdfColorGlyphArt.ToPathOperators(contours) : string.Empty;
            return new PdfColorGlyphArt(program.UnitsPerEm, new[] { new PdfColorGlyphLayer(path, OfficeColor.Black, foreground: true) });
        }

        private static bool HasAnyTable(byte[] data, params string[] tags) {
            if (data.Length < 12) return false;
            int count = data[4] << 8 | data[5];
            for (int index = 0; index < count && 12 + index * 16 + 4 <= data.Length; index++) {
                if (Array.IndexOf(tags, System.Text.Encoding.ASCII.GetString(data, 12 + index * 16, 4)) >= 0) return true;
            }
            return false;
        }
    }

    /// <summary>Whether the font has COLR layers or PNG glyphs, so its runs need glyph positions for their artwork.</summary>
    internal bool HasPaintedColorArt => _colorGlyphSource.Value.HasPaintedArt;

    /// <summary>
    /// Attaches the painted artwork of every glyph when the run shows a COLR or PNG color glyph. PDF text
    /// embeds outlines only, so the writer paints this artwork over the invisible, extractable text.
    /// </summary>
    private PdfTextShowCommand WithColorGlyphArt(PdfTextShowCommand command, PdfGlyphRun run) {
        ColorGlyphSource source = _colorGlyphSource.Value;
        if (!source.HasPaintedArt || run.Glyphs.Count == 0) return command;
        var art = new PdfColorGlyphArt?[run.Glyphs.Count];
        bool color = false;
        for (int index = 0; index < art.Length; index++) {
            art[index] = source.GetColorArt(run.Glyphs[index].GlyphId);
            color |= art[index] != null;
        }
        if (!color) return command;
        for (int index = 0; index < art.Length; index++) {
            art[index] ??= source.GetOutlineArt(run.Glyphs[index].GlyphId);
        }
        return command.WithColorArt(art);
    }

    /// <summary>
    /// Whether a painted glyph has no artwork the writer can draw: SVG or unreadable color data, or a PNG
    /// glyph in text written outside the page, such as a watermark Form. PDF text paints outlines only, so
    /// such a glyph would leave a silent blank. The blank glyphs of whitespace and nonpainting controls are
    /// intentional and excluded.
    /// </summary>
    internal bool IsUnpaintableColorGlyph(int glyphId) {
        ColorGlyphSource source = _colorGlyphSource.Value;
        if (source.Program == null || source.BlankGlyphs!.Contains(glyphId)) return false;
        return source.Program.IsUnpaintableColorGlyph(glyphId)
            || ContentStreamBuilder.BitmapGlyphPlacementSuppressed && source.Program.IsBitmapColorGlyph(glyphId);
    }

    /// <summary>Whether the font has color artwork that the preflight must inspect glyph by glyph.</summary>
    internal bool HasBitmapColorGlyphs {
        get {
            OfficeTrueTypeFont? program = _colorGlyphSource.Value.Program;
            return program != null && (program.HasBitmapGlyphs || program.HasUnreadColorTable);
        }
    }
}
