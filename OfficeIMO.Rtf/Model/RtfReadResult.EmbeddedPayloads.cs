using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfReadResult {
    /// <summary>Inventories inert pict and objdata destinations from the original immutable syntax, including semantically omitted objects.
    /// Locations are native RTF group character positions, not physical byte offsets or OPC relationships. Hexadecimal lengths remain
    /// unknown without a requested bounded read or hash. Semantic and lossless edits do not replace these original-source session bytes.</summary>
    public OfficeEmbeddedPayloadInspection InspectEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions? options = null,
        CancellationToken cancellationToken = default) => new OfficeEmbeddedPayloadInspection(GetSourcePayloads(cancellationToken), options, cancellationToken);

    /// <summary>Returns original-source payload descriptors under explicit limits. Optional hashes scan complete payloads within those limits.</summary>
    public IReadOnlyList<OfficeEmbeddedPayloadInfo> GetEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions options,
        CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        using OfficeEmbeddedPayloadInspection inspection = InspectEmbeddedPayloads(options, cancellationToken);
        return inspection.Payloads;
    }

    /// <summary>Reads one complete original-source payload under an explicit decoded byte ceiling without activating an object.</summary>
    public Stream OpenEmbeddedPayload(string id, long maximumBytes, CancellationToken cancellationToken = default) {
        if (id == null) throw new ArgumentNullException(nameof(id));
        if (maximumBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        foreach (OfficeEmbeddedPayloadSource source in GetSourcePayloads(cancellationToken)) {
            if (source.Info.Id != id) continue;
            using var inspection = new OfficeEmbeddedPayloadInspection(new[] { source }, new OfficeEmbeddedPayloadInspectionOptions {
                MaxPayloads = 1, MaxPayloadBytes = maximumBytes, MaxTotalPayloadBytes = maximumBytes, MaxReadBytes = maximumBytes
            }, cancellationToken);
            return inspection.OpenRead(0, cancellationToken);
        }
        throw new ArgumentException("Original-source payload id was not found.", nameof(id));
    }

    private IEnumerable<OfficeEmbeddedPayloadSource> GetSourcePayloads(CancellationToken token) {
        var pending = new Stack<RtfGroup>(); pending.Push(SyntaxTree.Root);
        while (pending.Count > 0) {
            token.ThrowIfCancellationRequested();
            RtfGroup group = pending.Pop();
            string? destination = group.Destination;
            if (destination == "pict" || destination == "objdata") {
                RtfPayloadSource snapshot = RtfPayloadSource.FromGroup(group, destination == "objdata", token.ThrowIfCancellationRequested, countHex: false);
                string owner = "rtf-source:/group/" + group.Position.ToString(CultureInfo.InvariantCulture);
                string id = owner + "/" + destination;
                string contentType = destination == "pict" && group.Children.OfType<RtfControlWord>().Any(control => control.Name == "pngblip")
                    ? "image/png" : destination == "pict" && group.Children.OfType<RtfControlWord>().Any(control => control.Name == "jpegblip")
                    ? "image/jpeg" : "application/octet-stream";
                var info = new OfficeEmbeddedPayloadInfo(id, destination == "objdata" ? OfficeEmbeddedPayloadKind.OleObject : OfficeEmbeddedPayloadKind.Other,
                    owner, string.Empty, id + "/data", contentType, "source-" + group.Position.ToString(CultureInfo.InvariantCulture) + ".bin",
                    snapshot.DeclaredLength, null);
                yield return new OfficeEmbeddedPayloadSource(info, snapshot.OpenRead);
            }
            for (int index = group.Children.Count - 1; index >= 0; index--)
                if (group.Children[index] is RtfGroup child) pending.Push(child);
        }
    }
}
