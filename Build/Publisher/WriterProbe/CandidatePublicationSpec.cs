using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Ordered physical pages, one master and their semantic objects, before native identities are assigned.</summary>
internal sealed record CandidatePublicationSpec(uint WidthEmu, uint HeightEmu, IReadOnlyList<CandidateObjectSpec> Objects,
    int PhysicalPageCount = 1);

internal enum CandidateObjectKind { TextBox, Rectangle, Picture, Table }

/// <summary>Placement in page points; the drawing codec alone converts it to native coordinates.</summary>
internal sealed record CandidatePlacement(double Left, double Top, double Width, double Height);

/// <summary>One owned object; PageIndex selects a physical page unless OnMaster selects the single master.</summary>
internal sealed record CandidateObjectSpec(CandidateObjectKind Kind, CandidatePlacement Placement,
    CandidateStorySpec? Story = null, byte[]? Image = null, int PageIndex = 0, double? ColumnSpacing = null,
    OfficeShape? Paint = null, string? ImageName = null, OfficeImageSourceCrop Crop = default, int? ImageId = null,
    bool OnMaster = false, uint FrameOrder = 0, PublisherTable? Table = null) {
    internal uint NativeKind => Kind == CandidateObjectKind.Picture ? 0x20U : Kind == CandidateObjectKind.Table ? 0x10U : 1U;
}

/// <summary>Owns a complete plain or styled story; table cell endpoints refer to that same text.</summary>
internal sealed class CandidateStorySpec {
    private readonly string? _plain;
    internal IReadOnlyList<OfficeRichTextParagraph>? Paragraphs { get; }
    internal IReadOnlyList<int>? CellEnds { get; }
    internal string Text => Paragraphs == null ? _plain! : string.Concat(Paragraphs.Select(paragraph =>
        string.Concat(paragraph.Runs.Select(run => run.Text)) + '\r'));

    private CandidateStorySpec(string? plain, IReadOnlyList<OfficeRichTextParagraph>? paragraphs, IReadOnlyList<int>? cellEnds = null) {
        _plain = plain; Paragraphs = paragraphs; CellEnds = cellEnds;
    }

    internal static CandidateStorySpec Plain(string text) {
        if (!text.EndsWith('\r')) text += '\r';
        if (text.Length > 4096) throw new ArgumentException("Probe text is limited to 4,096 characters.");
        return new CandidateStorySpec(text, null);
    }

    internal static CandidateStorySpec Styled(OfficeRichTextParagraph paragraph) => new(null, new[] { paragraph });

    internal static CandidateStorySpec Table(PublisherTable table, PublisherTextStory story) {
        int end = 0;
        int[] ends = table.Cells.Select(cell => end = checked(end + cell.Paragraphs.Sum(paragraph =>
            paragraph.Runs.Sum(run => run.Text.Length) + 1))).ToArray();
        var result = new CandidateStorySpec(null, story.Paragraphs, ends);
        if (result.Text != story.Text.Replace('\n', '\r') || ends.Last() != result.Text.Length)
            throw new NotSupportedException("The table specimen requires complete ordinary cell paragraphs.");
        return result;
    }
}
