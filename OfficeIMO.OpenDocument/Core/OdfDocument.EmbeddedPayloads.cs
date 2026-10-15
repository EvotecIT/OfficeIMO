namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    /// <summary>Inventories inert non-XML package payloads without expanding deferred entries unless hashing is requested.
    /// The session snapshots native entry identities and content sources; later edits do not replace its selected bytes.</summary>
    public OfficeEmbeddedPayloadInspection InspectEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions? options = null,
        CancellationToken cancellationToken = default) => new OfficeEmbeddedPayloadInspection(GetPayloadSources(cancellationToken), options, cancellationToken);

    /// <summary>Returns native non-XML payload descriptors under explicit inspection limits.</summary>
    public IReadOnlyList<OfficeEmbeddedPayloadInfo> GetEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions options,
        CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        using OfficeEmbeddedPayloadInspection inspection = InspectEmbeddedPayloads(options, cancellationToken);
        return inspection.Payloads;
    }

    /// <summary>Reads one complete inert non-XML native package payload under an explicit decoded byte ceiling.</summary>
    public Stream OpenEmbeddedPayload(string id, long maximumBytes, CancellationToken cancellationToken = default) {
        if (id == null) throw new ArgumentNullException(nameof(id));
        if (maximumBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        foreach (OfficeEmbeddedPayloadSource source in GetPayloadSources(cancellationToken)) {
            if (source.Info.Id != id) continue;
            using var inspection = new OfficeEmbeddedPayloadInspection(new[] { source }, new OfficeEmbeddedPayloadInspectionOptions {
                MaxPayloads = 1, MaxPayloadBytes = maximumBytes, MaxTotalPayloadBytes = maximumBytes, MaxReadBytes = maximumBytes
            }, cancellationToken);
            return inspection.OpenRead(0, cancellationToken);
        }
        throw new ArgumentException("Native embedded payload id was not found.", nameof(id));
    }

    private IEnumerable<OfficeEmbeddedPayloadSource> GetPayloadSources(CancellationToken token) {
        foreach (OdfPackageEntry current in Package.Entries) {
            token.ThrowIfCancellationRequested();
            if (!OdfPackage.IsBinaryPayloadPath(current.Name)) continue;
            OdfPackageEntry snapshot = current.CloneForSerialization(cloneXml: true);
            string uri = "/" + snapshot.Name;
            var info = new OfficeEmbeddedPayloadInfo("odf-entry:" + uri,
                snapshot.MediaType == "application/vnd.sun.star.oleobject" ? OfficeEmbeddedPayloadKind.OleObject : OfficeEmbeddedPayloadKind.Other,
                "/", string.Empty, uri, snapshot.MediaType ?? string.Empty, Path.GetFileName(snapshot.Name), snapshot.ByteLength, null);
            yield return new OfficeEmbeddedPayloadSource(info, snapshot.OpenRead);
        }
    }
}
