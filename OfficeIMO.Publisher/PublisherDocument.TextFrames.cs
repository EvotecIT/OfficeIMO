using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Replaces a qualified text frame's column count, gap, insets and vertical alignment, then recalculates current text ranges, overflow and page scenes. Dimensions and insets use points.</summary>
    /// <remarks>Columns must be between 1 and 256 within the configured item limit; gap is the full distance between adjacent content columns and must be nonnegative and finite. Insets and gaps must leave a positive finite content area. Top, Center and Bottom apply independently to generated content regions; native balancing is not provided. Complete stories, frame links, other frame settings, placement and graphics remain unchanged. Table stories, unresolved/transformed frames and unqualified native layout overrides are rejected. Existing references remain snapshots. Validation, cancellation or projection failure leaves the publication unchanged. Imported frames can be saved through the bounded native source profile when changed distances have exact nonnegative signed-EMU representations and their source records admit safe growth. Native save fidelity remains Unassessed and requires an explicit loss-policy choice.</remarks>
    public void SetTextFrameLayout(uint frameId, int columnCount, double columnSpacing, OfficeTextPadding padding,
        OfficeTextVerticalAlignment verticalAlignment, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (columnCount < 1 || columnCount > 256) throw new ArgumentOutOfRangeException(nameof(columnCount));
        if (!Finite(columnSpacing) || columnSpacing < 0) throw new ArgumentOutOfRangeException(nameof(columnSpacing));
        if (!Enum.IsDefined(typeof(OfficeTextVerticalAlignment), verticalAlignment)) throw new ArgumentOutOfRangeException(nameof(verticalAlignment));
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.SetFrameLayout(frameId, columnCount, columnSpacing, padding, verticalAlignment).Project(cancellationToken), cancellationToken);
    }

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
