using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Replaces a qualified text frame's column count, physical gap, insets and vertical alignment, then recalculates current text ranges, overflow and page scenes. Padding retains its drawing-owned point coordinates.</summary>
    /// <remarks>Columns must be between 1 and 256 within the configured item limit; gap is the full distance between adjacent content columns and must be nonnegative and finite. Insets and gaps must leave a positive finite content area. Top, Center and Bottom apply independently to generated content regions; native balancing is not provided. Complete stories, frame links, other frame settings, placement and graphics remain unchanged. Table stories, unresolved/transformed frames and unqualified native layout overrides are rejected. Existing references remain snapshots. Validation, cancellation or projection failure leaves the publication unchanged. Imported frames can be saved through the bounded native source profile when changed distances have exact nonnegative signed-EMU representations and their source records admit safe growth. Native save fidelity remains Unassessed and requires an explicit loss-policy choice.</remarks>
    public void SetTextFrameLayout(uint frameId, int columnCount, OfficeLength columnSpacing, OfficeTextPadding padding,
        OfficeTextVerticalAlignment verticalAlignment, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentOutOfRangeException.ThrowIfLessThan(columnCount, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(columnCount, 256);
        ArgumentOutOfRangeException.ThrowIfNegative(columnSpacing.ToEmus(), nameof(columnSpacing));
        if (!Enum.IsDefined(typeof(OfficeTextVerticalAlignment), verticalAlignment)) throw new ArgumentOutOfRangeException(nameof(verticalAlignment));
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.SetFrameLayout(frameId, columnCount, columnSpacing.ToPoints(), padding, verticalAlignment).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Appends a new continuation frame to an existing resolved non-table story, on an existing document page. Returns the current snapshot of the new frame.</summary>
    /// <remarks>The previous frame must be the tail of a complete chain of ordinary visible untransformed document-page frames. Bounds use physical page-local lengths; width and height must be positive. The frame shares the complete styled story, copies the preceding frame's columns, padding and vertical alignment, and has no fill, outline or wrap references. It is appended to the destination page's drawing order. Validation, cancellation or projection failure leaves the document unchanged. Existing snapshots remain detached; concurrent mutation is unsupported. This edits the in-memory publication and does not provide native PUB saving.</remarks>
    public PublisherTextFrame AddLinkedTextFrame(uint previousFrameId, uint pageId, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width.ToEmus(), nameof(width));
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height.ToEmus(), nameof(height));
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        PublisherDocument candidate = model.AddLinkedTextFrame(previousFrameId, pageId, x.ToPoints(), y.ToPoints(), width.ToPoints(), height.ToPoints(),
            cancellationToken, out uint frameId).Project(cancellationToken);
        PublisherTextFrame result = candidate.Pages.SelectMany(page => page.TextFrames).Single(frame => frame.Id == frameId);
        Commit(candidate, cancellationToken);
        return result;
    }
}
