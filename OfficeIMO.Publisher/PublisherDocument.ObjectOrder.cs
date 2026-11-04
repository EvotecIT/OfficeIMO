using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Sets an existing document or master page's back-to-front object order. Supply every identifier from its current <see cref="PublisherPage.ObjectIds"/> exactly once.</summary>
    /// <remarks>Identifiers, content, ownership and text-frame links are retained. Qualified rectangular wrap relationships are rebuilt in the model and native output. Grouped artwork and unqualified side, outline or master wrapping remain unsupported. Validation, cancellation or projection failure leaves the publication unchanged; previous page snapshots remain detached. Native saving requires a qualified imported source and the existing explicit loss policy.</remarks>
    public void SetPageObjectOrder(uint pageId, IReadOnlyList<uint> objectIds, CancellationToken cancellationToken = default) {
        if (objectIds == null) throw new ArgumentNullException(nameof(objectIds));
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        Commit(model.SetPageObjectOrder(pageId, objectIds, cancellationToken).Project(cancellationToken), cancellationToken);
    }
}
