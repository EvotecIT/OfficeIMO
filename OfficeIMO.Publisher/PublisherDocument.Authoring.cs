using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Creates an in-memory publication with one empty document page, no applied master and no native source. Dimensions use points and must be positive and finite.</summary>
    /// <remarks>Options are copied and bound current pages, objects, stories, assets and projection work. Source evidence starts at zero. This provides SVG/PDF rendering and supported semantic editing; it does not provide native PUB saving.</remarks>
    public static PublisherDocument Create(double width, double height, PublisherReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!Finite(width) || width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (!Finite(height) || height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        PublisherReadOptions operation = (options ?? new PublisherReadOptions()).Clone();
        return PublisherModelState.Create(width, height, operation, cancellationToken).Project(cancellationToken);
    }

    /// <summary>Appends an empty document page at the publication's existing size. A null master identifier means no applied master. Returns the committed immutable page snapshot.</summary>
    /// <remarks>Only qualified direct master definitions can be assigned. Existing page order, imported registration and object ownership stay intact. Page limits, cancellation or projection failure leave the publication and identifier allocator unchanged.</remarks>
    public PublisherPage AddPage(uint? masterPageId = null, CancellationToken cancellationToken = default) =>
        AddPageCore(false, string.Empty, masterPageId, cancellationToken);

    /// <summary>Appends an empty centered single-page master at the publication's existing size, returning its committed immutable snapshot. The name must contain non-control text.</summary>
    /// <remarks>The master has no inherited master. Populate it through supported shape, table and ordinary text-frame operations, then assign it to document pages. Use AddMasterPagePair for a left/right pair. Native PUB saving is not provided.</remarks>
    public PublisherPage AddMasterPage(string name, CancellationToken cancellationToken = default) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(name) || name.Any(char.IsControl))
            throw new ArgumentException("Master names must contain non-control text.", nameof(name));
        return AddPageCore(true, name, null, cancellationToken);
    }

    /// <summary>Assigns a qualified direct master to a document page, or clears the assignment when the master identifier is null. Inherited objects remain owned once by their master.</summary>
    /// <remarks>Qualified existing-source assignments can be saved natively. Clearing suppresses native master formatting while retaining its latent source reference; assigning a definition enables formatting. Master inheritance, unassessed backgrounds, native relationship tokens and registrations are rejected when unqualified. Dimensions and imported registration remain unchanged. Validation, cancellation or projection failure preserves the previous publication; existing snapshots remain detached.</remarks>
    public void SetPageMaster(uint pageId, uint? masterPageId, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.SetPageMaster(pageId, masterPageId).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Appends an independent styled text frame and complete story to a qualified document or master page. Bounds use page-local points; dimensions must be positive and finite. Returns the committed immutable frame snapshot.</summary>
    /// <remarks>The frame uses one column, zero gap and padding, top alignment, no fill or outline, and no transform or wrap references. Paragraphs are copied and carriage returns normalize to newlines; an empty collection creates an empty story. Ordinary authored stories have no implicit final terminator. Complete overflow remains in TextStories. Continue a document-page story with AddLinkedTextFrame. Limits, validation, cancellation or projection failure leave the publication and both reserved identifiers unchanged. Native PUB saving is not provided.</remarks>
    public PublisherTextFrame AddTextFrame(uint pageId, double x, double y, double width, double height,
        IReadOnlyList<OfficeRichTextParagraph> paragraphs, CancellationToken cancellationToken = default) {
        if (paragraphs == null) throw new ArgumentNullException(nameof(paragraphs));
        cancellationToken.ThrowIfCancellationRequested();
        if (!Finite(x)) throw new ArgumentOutOfRangeException(nameof(x));
        if (!Finite(y)) throw new ArgumentOutOfRangeException(nameof(y));
        if (!Finite(width) || width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (!Finite(height) || height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        PublisherDocument candidate = model.AddTextFrame(pageId, x, y, width, height, paragraphs, cancellationToken, out uint id).Project(cancellationToken);
        PublisherTextFrame result = candidate.Pages.Concat(candidate.MasterPages).SelectMany(page => page.TextFrames).Single(frame => frame.Id == id);
        Commit(candidate, cancellationToken);
        return result;
    }

    private PublisherPage AddPageCore(bool isMaster, string name, uint? masterId, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        PublisherDocument candidate = model.AddPage(isMaster, name, masterId, token, out uint id).Project(token);
        PublisherPage result = candidate.Pages.Concat(candidate.MasterPages).Single(page => page.Id == id);
        Commit(candidate, token);
        return result;
    }
}
