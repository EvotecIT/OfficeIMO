namespace OfficeIMO.Publisher.Internal;

/// <summary>One source-bound definition catalog. Ordinals and opaque metadata do not establish native named-style or inheritance semantics.</summary>
internal sealed class PublisherQuillStyleCatalog {
    internal static PublisherQuillStyleCatalog Empty { get; } = new(Array.Empty<PublisherQuillStyleDefinition>(),
        Array.Empty<PublisherQuillStyleLibrary>(), Array.Empty<string>(), Array.Empty<uint>());
    internal PublisherQuillStyleCatalog(IEnumerable<PublisherQuillStyleDefinition> definitions,
        IEnumerable<PublisherQuillStyleLibrary> libraries, IEnumerable<string> fonts, IEnumerable<uint> colors,
        IEnumerable<PublisherQuillStoryStyles>? stories = null) {
        Definitions = Array.AsReadOnly(definitions.ToArray()); Libraries = Array.AsReadOnly(libraries.ToArray());
        Fonts = Array.AsReadOnly(fonts.ToArray()); Colors = Array.AsReadOnly(colors.ToArray());
        Stories = Array.AsReadOnly((stories ?? Array.Empty<PublisherQuillStoryStyles>()).ToArray());
    }
    public IReadOnlyList<PublisherQuillStyleDefinition> Definitions { get; }
    public IReadOnlyList<PublisherQuillStyleLibrary> Libraries { get; }
    public IReadOnlyList<string> Fonts { get; }
    public IReadOnlyList<uint> Colors { get; }
    public IReadOnlyList<PublisherQuillStoryStyles> Stories { get; }
    internal PublisherQuillStyleCatalog WithStories(IEnumerable<PublisherQuillStoryStyles> stories) =>
        new(Definitions, Libraries, Fonts, Colors, stories);
}

/// <summary>A native STSH1 character/paragraph pair. Whole indexed entries retain uninterpreted prefixes and suffixes beside the decoded OPL spans.</summary>
internal sealed class PublisherQuillStyleDefinition {
    internal PublisherQuillStyleDefinition(int ordinal, PublisherCharacterStyle character, PublisherParagraphStyle paragraph,
        PublisherQuillSourceSpan characterEntry, PublisherQuillSourceSpan paragraphEntry,
        PublisherQuillSourceSpan characterDefinition, PublisherQuillSourceSpan paragraphDefinition) {
        Ordinal = ordinal; Character = character; Paragraph = paragraph;
        CharacterEntry = characterEntry; ParagraphEntry = paragraphEntry;
        CharacterDefinition = characterDefinition; ParagraphDefinition = paragraphDefinition;
    }
    public int Ordinal { get; }
    public PublisherCharacterStyle Character { get; }
    public PublisherParagraphStyle Paragraph { get; }
    public PublisherQuillSourceSpan CharacterEntry { get; }
    public PublisherQuillSourceSpan ParagraphEntry { get; }
    public PublisherQuillSourceSpan CharacterDefinition { get; }
    public PublisherQuillSourceSpan ParagraphDefinition { get; }
}

/// <summary>Indexed STSH source envelopes. STSH0/2 entries remain opaque; their ordinals are not interpreted as names, parents or following styles.</summary>
internal sealed class PublisherQuillStyleLibrary {
    internal PublisherQuillStyleLibrary(int id, PublisherQuillSourceSpan source, IEnumerable<PublisherQuillSourceSpan> entries) {
        Id = id; Source = source; Entries = Array.AsReadOnly(entries.ToArray());
    }
    public int Id { get; }
    public PublisherQuillSourceSpan Source { get; }
    public IReadOnlyList<PublisherQuillSourceSpan> Entries { get; }
}

/// <summary>Byte coordinates in Quill/QuillSub/CONTENTS, independent of normalized public story text.</summary>
internal readonly record struct PublisherQuillSourceSpan(int Offset, int Length);

/// <summary>An original formatting-range intersection with one story. NativeStart/Length count UTF-16 units from TEXT start, including CR and NUL.</summary>
internal readonly record struct PublisherQuillStyleUse(string Tag, int ChunkId, int RangeIndex,
    int NativeStart, int NativeLength, PublisherQuillSourceSpan DirectDefinition, uint? SelectedStyleOrdinal);

/// <summary>Original story identity, native character domain and style dependencies. A null FDPP selector means absent, even when resolution uses pair zero.</summary>
internal sealed class PublisherQuillStoryStyles {
    internal PublisherQuillStoryStyles(uint storyId, int storyOrdinal, int nativeStart, int nativeLength, int textByteOffset,
        int carriageReturns, int nullCharacters, IEnumerable<PublisherQuillStyleUse> uses,
        IEnumerable<PublisherQuillStyleDefinition> definitions) {
        StoryId = storyId; StoryOrdinal = storyOrdinal; NativeStart = nativeStart; NativeLength = nativeLength;
        TextByteOffset = textByteOffset; CarriageReturns = carriageReturns; NullCharacters = nullCharacters;
        Uses = Array.AsReadOnly(uses.ToArray()); Definitions = Array.AsReadOnly(definitions.ToArray());
    }
    public uint StoryId { get; }
    public int StoryOrdinal { get; }
    public int NativeStart { get; }
    public int NativeLength { get; }
    public int TextByteOffset { get; }
    public int CarriageReturns { get; }
    public int NullCharacters { get; }
    public IReadOnlyList<PublisherQuillStyleUse> Uses { get; }
    public IReadOnlyList<PublisherQuillStyleDefinition> Definitions { get; }
}
