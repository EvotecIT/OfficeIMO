using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Appends both definitions of a two-page master atomically and returns their committed immutable snapshots.</summary>
    /// <remarks>Both definitions have the supplied name and the publication's existing size. Populate and assign each member through the ordinary page APIs. Definition coordinates use the native left/right registration: an applied left definition shifts right by half a page width, and an applied right definition shifts left by half a page width. No master is selected automatically by page parity. Limits, invalid input, cancellation or projection failure leave both definitions and the identifier allocator unchanged. This in-memory operation does not provide native PUB saving.</remarks>
    public (PublisherPage Left, PublisherPage Right) AddMasterPagePair(string name, CancellationToken cancellationToken = default) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(name) || name.Any(char.IsControl))
            throw new ArgumentException("Master names must contain non-control text.", nameof(name));
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable model.");
        PublisherDocument candidate = model.AddMasterPagePair(name, cancellationToken, out uint leftId, out uint rightId).Project(cancellationToken);
        PublisherPage left = candidate.MasterPages.Single(page => page.Id == leftId);
        PublisherPage right = candidate.MasterPages.Single(page => page.Id == rightId);
        Commit(candidate, cancellationToken);
        return (left, right);
    }
}
