using OfficeIMO.OpenXml.Internal;
using System.Threading;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    /// <summary>Inventories the VBA part and signature relationship without reading its bytes or decoding module names. Payload length remains unknown.</summary>
    public OfficeVbaProjectInfo? InspectVbaProjectMetadata() {
        ThrowIfDisposed();
        return _presentationPart.VbaProjectPart is { } part ? OfficeOpenXmlPackagePayload.CreateVbaDescriptor(part) : null;
    }

    /// <summary>Inventories inert native payloads and creates a bounded cumulative read session. Keep this presentation open for new reads; returned streams are detached and caller-owned.</summary>
    public OfficeEmbeddedPayloadInspection InspectEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions? options = null, CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        return OfficeOpenXmlPackagePayload.Inspect(OpenXmlDocument, options, cancellationToken);
    }

    /// <summary>Inventories payloads within explicit count/member/aggregate/work ceilings without scanning unrequested bytes unless hashing is requested.</summary>
    public IReadOnlyList<OfficeEmbeddedPayloadInfo> GetEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions options, CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        using OfficeEmbeddedPayloadInspection inspection = InspectEmbeddedPayloads(options, cancellationToken);
        return inspection.Payloads;
    }

    /// <summary>Opens a complete bounded read-only detached payload stream. The caller owns its lifetime; no external target is accessed.</summary>
    public Stream OpenEmbeddedPayload(string id, long maximumBytes, CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        return OfficeOpenXmlPackagePayload.OpenDetached(OpenXmlDocument, id, maximumBytes, cancellationToken);
    }
}
