using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Ordered physical pages, master definitions and their semantic objects, before native identities are assigned.</summary>
internal sealed record PublisherNativeCreationPlan(uint WidthEmu, uint HeightEmu, IReadOnlyList<PublisherNativeCreationObject> Objects,
    int PhysicalPageCount = 1, IReadOnlyList<PublisherNativeCreationMaster>? Masters = null, IReadOnlyList<int?>? PageMasters = null) {
    private static readonly PublisherNativeCreationMaster[] DefaultMasters = { new("A", "Master Page A", PublisherMasterPageSide.Single) };
    internal IReadOnlyList<PublisherNativeCreationMaster> MasterDefinitions => Masters ?? DefaultMasters;
}

/// <summary>A named native master definition; left/right members retain their individual content and geometry.</summary>
internal sealed record PublisherNativeCreationMaster(string Code, string Name, PublisherMasterPageSide Side);

internal enum PublisherNativeCreationObjectKind { TextBox, Rectangle, Picture, Table }

/// <summary>Placement in page points; the drawing codec alone converts it to native coordinates.</summary>
internal sealed record PublisherNativeCreationPlacement(double Left, double Top, double Width, double Height);

/// <summary>One owned object; PageIndex selects a physical page or, when OnMaster is true, its master definition.</summary>
internal sealed record PublisherNativeCreationObject(PublisherNativeCreationObjectKind Kind, PublisherNativeCreationPlacement Placement,
    PublisherNativeCreationStory? Story = null, byte[]? Image = null, int PageIndex = 0, double? ColumnSpacing = null,
    OfficeShape? Paint = null, string? ImageName = null, OfficeImageSourceCrop Crop = default, int? ImageId = null,
    bool OnMaster = false, uint FrameOrder = 0, PublisherTable? Table = null, PublisherTextFrame? FrameLayout = null) {
    internal uint NativeKind => Kind == PublisherNativeCreationObjectKind.Picture ? 0x20U : Kind == PublisherNativeCreationObjectKind.Table ? 0x10U : 1U;
}

/// <summary>Owns a complete plain or styled story; table cell endpoints refer to that same text.</summary>
internal sealed class PublisherNativeCreationStory {
    private readonly string? _plain;
    internal IReadOnlyList<OfficeRichTextParagraph>? Paragraphs { get; }
    internal IReadOnlyList<int>? CellEnds { get; }
    internal bool? AutomaticHyphenation { get; }
    /// <summary>Semantic story and label characters, retaining the canonical model's final-terminator convention.</summary>
    internal long CharacterCount { get; }
    internal string Text => Paragraphs == null ? _plain! : string.Concat(Paragraphs.Select(paragraph =>
        string.Concat(paragraph.Runs.Select(run => run.Text)) + '\r'));

    private PublisherNativeCreationStory(string? plain, IReadOnlyList<OfficeRichTextParagraph>? paragraphs, IReadOnlyList<int>? cellEnds = null,
        bool? automaticHyphenation = null, long characterCount = 0) {
        _plain = plain; Paragraphs = paragraphs; CellEnds = cellEnds; AutomaticHyphenation = automaticHyphenation;
        CharacterCount = characterCount;
    }

    internal static PublisherNativeCreationStory Plain(string text) {
        if (!text.EndsWith("\r", StringComparison.Ordinal)) text += '\r';
        if (text.Length > 4096) throw new ArgumentException("Probe text is limited to 4,096 characters.");
        return new PublisherNativeCreationStory(text, null, characterCount: text.Length);
    }

    internal static PublisherNativeCreationStory Styled(OfficeRichTextParagraph paragraph) => new(null, new[] { paragraph },
        characterCount: paragraph.Runs.Sum(run => (long)run.Text.Length) + (paragraph.Label?.Run.Text.Length ?? 0));

    internal static PublisherNativeCreationStory Styled(PublisherTextStory story) {
        if (story.Paragraphs.Count == 0)
            throw new NotSupportedException("The text specimen requires at least one complete paragraph per story.");
        if (story.Paragraphs.Any(paragraph => paragraph.Runs.Count == 0))
            throw new NotSupportedException("The text specimen requires at least one explicitly styled run per paragraph.");
        return new PublisherNativeCreationStory(null, story.Paragraphs, automaticHyphenation: story.AutomaticHyphenation,
            characterCount: PublisherStoryEditor.CharacterCount(story));
    }

    internal static PublisherNativeCreationStory Table(PublisherTable table, PublisherTextStory story) {
        int end = 0;
        int[] ends = table.Cells.Select(cell => end = checked(end + cell.Paragraphs.Sum(paragraph =>
            paragraph.Runs.Sum(run => run.Text.Length) + 1))).ToArray();
        var result = new PublisherNativeCreationStory(null, story.Paragraphs, ends, story.AutomaticHyphenation,
            PublisherStoryEditor.CharacterCount(story));
        if (result.Text != story.Text.Replace('\n', '\r') || ends.Last() != result.Text.Length)
            throw new NotSupportedException("The table specimen requires complete ordinary cell paragraphs.");
        return result;
    }
}
