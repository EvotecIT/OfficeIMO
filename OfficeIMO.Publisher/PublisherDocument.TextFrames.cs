using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Appends a new continuation frame to an existing resolved non-table story, on an existing document page. Returns the current snapshot of the new frame.</summary>
    /// <remarks>The previous frame must be the tail of a complete chain of ordinary visible untransformed document-page frames. Bounds use page-local points; width and height must be positive and finite. The frame shares the complete styled story, copies the preceding frame's columns, padding and vertical alignment, and has no fill, outline or wrap references. It is appended to the destination page's drawing order. Validation, cancellation or projection failure leaves the document unchanged. Existing snapshots remain detached; concurrent mutation is unsupported. This edits the in-memory publication and does not provide native PUB saving.</remarks>
    public PublisherTextFrame AddLinkedTextFrame(uint previousFrameId, uint pageId, double x, double y, double width, double height,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!Finite(x)) throw new ArgumentOutOfRangeException(nameof(x));
        if (!Finite(y)) throw new ArgumentOutOfRangeException(nameof(y));
        if (!Finite(width) || width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (!Finite(height) || height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        PublisherDocument candidate = model.AddLinkedTextFrame(previousFrameId, pageId, x, y, width, height,
            cancellationToken, out uint frameId).Project(cancellationToken);
        PublisherTextFrame result = candidate.Pages.SelectMany(page => page.TextFrames).Single(frame => frame.Id == frameId);
        Commit(candidate, cancellationToken);
        return result;
    }
}
