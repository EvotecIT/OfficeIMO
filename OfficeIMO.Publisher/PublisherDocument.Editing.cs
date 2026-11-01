using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    internal void AttachModel(PublisherModelState model, PublisherReadReport sourceReport) { _model = model; SourceReadReport = sourceReport; }

    /// <summary>Sets an ordinary story's native automatic-hyphenation flag without changing its source text or formatting.</summary>
    /// <remarks>Managed layout does not yet generate automatic hyphens. Native source saving requires a qualified existing story-catalog entry. Table stories are unsupported. Failed edits leave the publication unchanged; fetch current snapshots after an edit.</remarks>
    public void SetStoryAutomaticHyphenation(uint storyId, bool enabled, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        Commit(model.SetStoryAutomaticHyphenation(storyId, enabled).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Replaces a non-table story's styled paragraphs and recalculates its page layout. IDs and frame links are retained. Fetch current pages, stories and frames after an edit; previously returned objects remain snapshots.</summary>
    /// <remarks>Carriage returns in run text normalize to newlines. The story retains its existing final paragraph-terminator convention. A retained empty paragraph can save native CR typing: no runs retain the source typing style, while an empty styled run requests that style. Original empty paragraph snapshots retain their complete native CR definition when reordered or duplicated; copied values require unambiguous native definitions. Native reload exposes the styled empty run. Count-changing native edits resolve exact snapshots and values first, then retained ordered run text, complete paragraph text with preserved fields at each position, then unchanged complete formatting. Correspondence must be unambiguous across complete native templates. Retained-text run splitting/merging maps source, requested and native character ranges while preserving opaque fields and destination CR typing. Reload may expose separate native ranges with identical visible formatting. Changed-text replacement requires a uniform native definition for each corresponding ordered run. A single empty source run can supply its complete native CR typing for a nonempty replacement with unchanged formatting. Formatting edits use the qualified character/paragraph codecs; requests changing both text and formatting without correspondence refuse saving. Removing every paragraph remains a separate native-source limitation. Validation, cancellation or projection failure leaves the publication unchanged. Concurrent mutation is not supported.</remarks>
    public void ReplaceStoryParagraphs(uint storyId, IReadOnlyList<OfficeRichTextParagraph> paragraphs, CancellationToken cancellationToken = default) {
        if (paragraphs == null) throw new ArgumentNullException(nameof(paragraphs));
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        Commit(model.ReplaceStory(storyId, paragraphs, cancellationToken).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Sets a visible, resolved, untransformed top-level text frame's bounds in page-local points and recalculates the publication. Negative page offsets allow clipped bleed; width and height must be positive.</summary>
    /// <remarks>Grouped, rotated, reflected, unresolved and table frames are unsupported. IDs, story ownership and links are retained. Fetch current frame snapshots after an edit. Failed edits leave the publication unchanged; concurrent mutation is not supported.</remarks>
    public void SetTextFrameBounds(uint frameId, double x, double y, double width, double height, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!Finite(x)) throw new ArgumentOutOfRangeException(nameof(x));
        if (!Finite(y)) throw new ArgumentOutOfRangeException(nameof(y));
        if (!Finite(width) || width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (!Finite(height) || height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        PublisherTextFrame? frame = Pages.Concat(MasterPages).SelectMany(page => page.TextFrames).SingleOrDefault(item => item.Id == frameId);
        if (frame == null || !frame.TextStart.HasValue)
            throw new NotSupportedException("The object is not a text frame with resolved story placement.");
        Commit(model.SetBounds(frame, x, y, width, height).Project(cancellationToken), cancellationToken);
    }

    private void Commit(PublisherDocument candidate, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        Pages = candidate.Pages; MasterPages = candidate.MasterPages;
        TextStories = candidate.TextStories; Images = candidate.Images; ReadReport = candidate.ReadReport; SourceReadReport = candidate.SourceReadReport;
        _model = candidate._model;
    }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
