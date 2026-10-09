using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed class PublisherQuillStyles {
    private readonly PublisherParseContext _context;
    private readonly IReadOnlyList<OfficeColor?> _palette;
    internal PublisherQuillStyles(PublisherParseContext context, IReadOnlyList<OfficeColor?> palette) { _context = context; _palette = palette; }
    internal List<string> Fonts { get; } = new();
    internal List<uint> Colors { get; } = new();
    internal List<PublisherCharacterRange> Characters { get; } = new();
    internal List<PublisherParagraphRange> Paragraphs { get; } = new();
    internal PublisherCharacterStyle DefaultCharacter { get; set; } = new();
    internal PublisherParagraphStyle DefaultParagraph { get; set; } = new();
    internal PublisherCharacterRange? CharacterAt(int position) {
        int index = Find(Characters.Count, i => Characters[i].End, position);
        return index < Characters.Count ? Characters[index] : null;
    }
    internal PublisherParagraphStyle ParagraphAt(int position) {
        int index = Find(Paragraphs.Count, i => Paragraphs[i].End, position);
        return index < Paragraphs.Count ? Paragraphs[index].Style : DefaultParagraph;
    }
    internal OfficeRichTextRun Run(string text, PublisherCharacterStyle? style) {
        style ??= DefaultCharacter;
        double size = style.Size ?? DefaultCharacter.Size ?? 10;
        uint font = style.Font ?? DefaultCharacter.Font ?? 0;
        string family = font < Fonts.Count ? Fonts[(int)font] : "Times New Roman";
        if (font >= Fonts.Count) _context.Add("PUB_FONT_REFERENCE_UNRESOLVED", "A native font reference could not be resolved; the text uses Times New Roman.", OfficeConversionLossKind.Approximation, "Quill/FONT");
        uint? colorIndex = style.Color ?? DefaultCharacter.Color;
        OfficeColor color = OfficeColor.Black;
        if (colorIndex.HasValue && colorIndex.Value < Colors.Count) {
            var reference = new OfficeArtColorReference(Colors[(int)colorIndex.Value]);
            if (!reference.TryResolve(index => index < _palette.Count ? _palette[index] : null, out color)) {
                color = OfficeColor.Black;
                _context.Add("PUB_TEXT_COLOR_UNRESOLVED", "A native text color was approximated as black.", OfficeConversionLossKind.Approximation, "Quill/PL");
            }
        } else if (colorIndex.HasValue) _context.Add("PUB_TEXT_COLOR_UNRESOLVED", "A native text color index is unresolved; the text uses black.", OfficeConversionLossKind.Approximation, "Quill/PL");
        return new OfficeRichTextRun(text, size, color, style.Bold ?? DefaultCharacter.Bold ?? false,
            style.Italic ?? DefaultCharacter.Italic ?? false, style.Underline ?? DefaultCharacter.Underline ?? false,
            family, baseline: style.Baseline ?? DefaultCharacter.Baseline ?? OfficeTextBaseline.Normal);
    }
    private static int Find(int count, Func<int, int> end, int position) {
        int low = 0, high = count;
        while (low < high) { int mid = low + (high - low) / 2; if (end(mid) <= position) low = mid + 1; else high = mid; }
        return low;
    }
}

internal sealed class PublisherCharacterStyle {
    internal double? Size { get; set; }
    internal uint? Font { get; set; }
    internal uint? Color { get; set; }
    internal bool? Bold { get; set; }
    internal bool? Italic { get; set; }
    internal bool? Underline { get; set; }
    internal OfficeTextBaseline? Baseline { get; set; }
}
internal sealed class PublisherParagraphStyle {
    internal OfficeTextAlignment Alignment { get; set; }
    internal double? LineHeight { get; set; }
    internal double? LineHeightFactor { get; set; }
    internal OfficeTextPadding Margins { get; set; }
    internal OfficeTextParagraphIndent Indent { get; set; }
}
internal sealed class PublisherCharacterRange {
    internal PublisherCharacterRange(int end, PublisherCharacterStyle style) { End = end; Style = style; }
    internal int End { get; }
    internal PublisherCharacterStyle Style { get; }
}
internal sealed class PublisherParagraphRange {
    internal PublisherParagraphRange(int end, PublisherParagraphStyle style) { End = end; Style = style; }
    internal int End { get; }
    internal PublisherParagraphStyle Style { get; }
}
