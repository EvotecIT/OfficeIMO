using System.Threading;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

public sealed partial class IWorkNativeEditSession {
    /// <summary>Replaces one plain body paragraph while preserving its native text and byte offsets.</summary>
    /// <remarks>The paragraph number is one-based and the identity must match the source body text storage.
    /// Replacement text must retain the paragraph's UTF-16 length and UTF-8 byte length. Only complete, single-section
    /// body-only Pages sources and uniform, non-list paragraphs without hyperlinks or inline content are supported.
    /// Styles, delimiters, IDs, references and unknown bytes remain intact. Edited saves require explicit visual-cache acceptance.</remarks>
    public void SetPagesBodyParagraphText(IWorkObjectIdentity bodyIdentity, int paragraph, string text,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (bodyIdentity == null) throw new ArgumentNullException(nameof(bodyIdentity));
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (paragraph < 1) throw new ArgumentOutOfRangeException(nameof(paragraph));
        if (_source.Kind != IWorkDocumentKind.Pages) throw new InvalidOperationException("Paragraph editing requires a Pages source.");
        if (_source.ContainerEntries.Any(entry => entry.Path == "Index.zip"))
            throw new NotSupportedException("Edits to nested Index.zip sources are not supported.");
        ValidateParagraphReplacementText(text);

        var view = _source.WithCancellation(cancellationToken);
        var projection = view.ReadPages();
        var body = projection.Body;
        if (!projection.HasEditableContent || !body.IsComplete || projection.SourceDeclarationIssues.Count != 0
            || projection.SourceReferenceIssues.Count != 0 || projection.Sections.Count > 1 || projection.Drawables.Count != 0
            || projection.HeaderContents.Concat(projection.FooterContents).Any(content =>
                content.Paragraphs.Any(item => item.Text.Length != 0 || item.Runs.Any(run => run.InlineObject != null)))
            || body.Paragraphs.Any(item => item.Runs.Any(run => run.InlineObject != null)))
            throw new NotSupportedException("Paragraph editing requires complete single-section body-only Pages content.");
        if (!SameIdentity(body.SourceIdentity, bodyIdentity))
            throw new ArgumentException("The identity must resolve to this source's body text storage.", nameof(bodyIdentity));
        ReplaceNativeParagraphText(view, body, paragraph, text, cancellationToken);
    }
}
