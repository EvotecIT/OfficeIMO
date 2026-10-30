using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>One page and its ordered semantic objects, before native identities are assigned.</summary>
internal sealed record CandidatePublicationSpec(uint WidthEmu, uint HeightEmu, IReadOnlyList<CandidateObjectSpec> Objects);

internal enum CandidateObjectKind { TextBox, Rectangle, Picture }

/// <summary>Placement in page points; the drawing codec alone converts it to native coordinates.</summary>
internal sealed record CandidatePlacement(double Left, double Top, double Width, double Height);

internal sealed record CandidateObjectSpec(CandidateObjectKind Kind, CandidatePlacement Placement,
    CandidateStorySpec? Story = null, byte[]? Image = null) {
    internal uint NativeKind => Kind == CandidateObjectKind.Picture ? 0x20U : 1U;
}

/// <summary>Owns either plain text or a shared styled paragraph, with one native paragraph terminator.</summary>
internal sealed class CandidateStorySpec {
    private readonly string? _plain;
    internal OfficeRichTextParagraph? Paragraph { get; }
    internal string Text => Paragraph == null ? _plain! : string.Concat(Paragraph.Runs.Select(run => run.Text)) + '\r';

    private CandidateStorySpec(string? plain, OfficeRichTextParagraph? paragraph) { _plain = plain; Paragraph = paragraph; }

    internal static CandidateStorySpec Plain(string text) {
        if (!text.EndsWith('\r')) text += '\r';
        if (text.Length > 4096) throw new ArgumentException("Probe text is limited to 4,096 characters.");
        return new CandidateStorySpec(text, null);
    }

    internal static CandidateStorySpec Styled(OfficeRichTextParagraph paragraph) => new(null, paragraph);
}
