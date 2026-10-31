using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed class PublisherQuillStyles {
    private readonly PublisherParseContext _context;
    private readonly IReadOnlyList<OfficeColor?> _palette;
    internal PublisherQuillStyles(PublisherParseContext context, IReadOnlyList<OfficeColor?> palette) { _context = context; _palette = palette; }
    internal PublisherQuillStyleCatalog Catalog { get; set; } = PublisherQuillStyleCatalog.Empty;
    internal List<PublisherCharacterRange> Characters { get; } = new();
    internal List<PublisherParagraphRange> Paragraphs { get; } = new();
    private readonly PublisherCharacterStyle _emptyCharacter = new();
    private readonly PublisherParagraphStyle _emptyParagraph = new();
    private PublisherCharacterStyle DefaultCharacter => Catalog.Definitions.Count == 0 ? _emptyCharacter : Catalog.Definitions[0].Character;
    private PublisherParagraphStyle DefaultParagraph => Catalog.Definitions.Count == 0 ? _emptyParagraph : Catalog.Definitions[0].Paragraph;
    internal PublisherCharacterRange? CharacterAt(int position) {
        int index = Find(Characters.Count, i => Characters[i].End, position);
        return index < Characters.Count ? Characters[index] : null;
    }
    internal PublisherParagraphStyle ParagraphAt(int position) {
        int index = Find(Paragraphs.Count, i => Paragraphs[i].End, position);
        PublisherParagraphStyle direct = index < Paragraphs.Count ? Paragraphs[index].Style : DefaultParagraph;
        uint selected = direct.DefaultStyleIndex ?? 0;
        if (selected >= Catalog.Definitions.Count) {
            if (direct.DefaultStyleIndex.HasValue) _context.Add("PUB_TEXT_STYLE_REFERENCE_UNRESOLVED",
                "A paragraph references an unavailable native style. Recovered global defaults fill its unspecified values.", OfficeConversionLossKind.Approximation, "Quill/STSH");
            return direct.Inherit(DefaultParagraph);
        }
        return direct.Inherit(Catalog.Definitions[(int)selected].Paragraph).Inherit(DefaultParagraph);
    }
    internal OfficeRichTextRun Run(string text, PublisherCharacterStyle? style, PublisherParagraphStyle? paragraph = null) {
        uint selected = paragraph?.DefaultStyleIndex ?? 0;
        PublisherCharacterStyle fallback = selected < Catalog.Definitions.Count ? Catalog.Definitions[(int)selected].Character : DefaultCharacter;
        style = (style ?? new PublisherCharacterStyle()).Inherit(fallback).Inherit(DefaultCharacter);
        double size = style.Size ?? DefaultCharacter.Size ?? 10;
        uint font = style.Font ?? DefaultCharacter.Font ?? 0;
        string family = font < Catalog.Fonts.Count ? Catalog.Fonts[(int)font] : "Times New Roman";
        if (font >= Catalog.Fonts.Count) _context.Add("PUB_FONT_REFERENCE_UNRESOLVED", "A native font reference could not be resolved; the text uses Times New Roman.", OfficeConversionLossKind.Approximation, "Quill/FONT");
        uint? colorIndex = style.Color ?? DefaultCharacter.Color;
        OfficeColor color = OfficeColor.Black;
        if (colorIndex.HasValue && colorIndex.Value < Catalog.Colors.Count) {
            var reference = new OfficeArtColorReference(Catalog.Colors[(int)colorIndex.Value]);
            if (!reference.TryResolve(index => index < _palette.Count ? _palette[index] : null, out color)) {
                color = OfficeColor.Black;
                _context.Add("PUB_TEXT_COLOR_UNRESOLVED", "A native text color was approximated as black.", OfficeConversionLossKind.Approximation, "Quill/PL");
            }
        } else if (colorIndex.HasValue) _context.Add("PUB_TEXT_COLOR_UNRESOLVED", "A native text color index is unresolved; the text uses black.", OfficeConversionLossKind.Approximation, "Quill/PL");
        return new OfficeRichTextRun(text, size, color, style.Bold ?? DefaultCharacter.Bold ?? false,
            style.Italic ?? DefaultCharacter.Italic ?? false, fontFamily: family,
            strikethrough: style.Strikethrough ?? DefaultCharacter.Strikethrough ?? false,
            underlineStyle: style.UnderlineStyle ?? DefaultCharacter.UnderlineStyle ?? OfficeTextDecorationStyle.None,
            baseline: style.Baseline ?? DefaultCharacter.Baseline ?? OfficeTextBaseline.Normal) { LanguageId = style.LanguageId };
    }
    internal PublisherQuillStoryStyles BindStory(uint id, int ordinal, int start, string raw, int textOffset) {
        int end = checked(start + raw.Length);
        var uses = new List<PublisherQuillStyleUse>();
        var selected = new HashSet<int>();
        // Pair zero remains the resolver's global fallback, even when a range
        // explicitly selects another pair. It is not an encoded selector.
        if (raw.Length != 0 && Catalog.Definitions.Count != 0) selected.Add(0);
        int first = Find(Paragraphs.Count, index => Paragraphs[index].End, start);
        for (int index = first; index < Paragraphs.Count; index++) {
            int rangeStart = index == 0 ? 0 : Paragraphs[index - 1].End;
            if (rangeStart >= end) break;
            PublisherParagraphRange range = Paragraphs[index];
            _context.Record();
            int from = Math.Max(start, rangeStart), to = Math.Min(end, range.End);
            if (to <= from) continue;
            uses.Add(new PublisherQuillStyleUse("FDPP", range.ChunkId, range.RangeIndex, from, to - from, range.Source, range.Style.DefaultStyleIndex));
            uint choice = range.Style.DefaultStyleIndex ?? 0;
            if (Catalog.Definitions.Count != 0) selected.Add(choice < Catalog.Definitions.Count ? (int)choice : 0);
        }
        first = Find(Characters.Count, index => Characters[index].End, start);
        for (int index = first; index < Characters.Count; index++) {
            int rangeStart = index == 0 ? 0 : Characters[index - 1].End;
            if (rangeStart >= end) break;
            PublisherCharacterRange range = Characters[index];
            _context.Record();
            int from = Math.Max(start, rangeStart), to = Math.Min(end, range.End);
            if (to > from) uses.Add(new PublisherQuillStyleUse("FDPC", range.ChunkId, range.RangeIndex, from, to - from, range.Source, null));
        }
        return new PublisherQuillStoryStyles(id, ordinal, start, raw.Length, checked(textOffset + start * 2),
            raw.Count(character => character == '\r'), raw.Count(character => character == '\0'), uses,
            selected.OrderBy(ordinal => ordinal).Select(ordinal => Catalog.Definitions[ordinal]));
    }
    private static int Find(int count, Func<int, int> end, int position) {
        int low = 0, high = count;
        while (low < high) { int mid = low + (high - low) / 2; if (end(mid) <= position) low = mid + 1; else high = mid; }
        return low;
    }
}

internal sealed class PublisherCharacterStyle {
    public uint? LanguageId { get; internal set; }
    public double? Size { get; internal set; }
    public uint? Font { get; internal set; }
    public uint? Color { get; internal set; }
    public bool? Bold { get; internal set; }
    public bool? Italic { get; internal set; }
    public bool? Strikethrough { get; internal set; }
    public OfficeTextDecorationStyle? UnderlineStyle { get; internal set; }
    public OfficeTextBaseline? Baseline { get; internal set; }
    internal PublisherCharacterStyle Inherit(PublisherCharacterStyle fallback) => new() {
        Size = Size ?? fallback.Size, Font = Font ?? fallback.Font, Color = Color ?? fallback.Color,
        Bold = Bold ?? fallback.Bold, Italic = Italic ?? fallback.Italic, Strikethrough = Strikethrough ?? fallback.Strikethrough,
        UnderlineStyle = UnderlineStyle ?? fallback.UnderlineStyle,
        Baseline = Baseline ?? fallback.Baseline, LanguageId = LanguageId ?? fallback.LanguageId
    };
}
internal sealed class PublisherCharacterRange {
    internal PublisherCharacterRange(int end, PublisherCharacterStyle style, PublisherQuillSourceSpan source, int chunkId, int rangeIndex) {
        End = end; Style = style; Source = source; ChunkId = chunkId; RangeIndex = rangeIndex;
    }
    internal int End { get; }
    internal PublisherCharacterStyle Style { get; }
    internal PublisherQuillSourceSpan Source { get; }
    internal int ChunkId { get; }
    internal int RangeIndex { get; }
}
internal sealed class PublisherParagraphRange {
    internal PublisherParagraphRange(int end, PublisherParagraphStyle style, PublisherQuillSourceSpan source, int chunkId, int rangeIndex) {
        End = end; Style = style; Source = source; ChunkId = chunkId; RangeIndex = rangeIndex;
    }
    internal int End { get; }
    internal PublisherParagraphStyle Style { get; }
    internal PublisherQuillSourceSpan Source { get; }
    internal int ChunkId { get; }
    internal int RangeIndex { get; }
}
