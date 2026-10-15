using OfficeIMO.OpenXml.Internal;
using System.Threading;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Inventories the VBA part and signature relationship without reading its bytes or decoding module names. Payload length remains unknown.</summary>
    public OfficeVbaProjectInfo? InspectVbaProjectMetadata() => _wordprocessingDocument.MainDocumentPart?.VbaProjectPart is { } part
        ? OfficeOpenXmlPackagePayload.CreateVbaDescriptor(part) : null;

    /// <summary>Inventories inert native payloads and creates a bounded cumulative read session. Keep this document open for new reads; returned streams are detached and caller-owned.</summary>
    public OfficeEmbeddedPayloadInspection InspectEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeOpenXmlPackagePayload.Inspect(_wordprocessingDocument, options, cancellationToken);

    /// <summary>Inventories payloads within explicit count/member/aggregate/work ceilings without scanning unrequested bytes unless hashing is requested.</summary>
    public IReadOnlyList<OfficeEmbeddedPayloadInfo> GetEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions options, CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        using OfficeEmbeddedPayloadInspection inspection = InspectEmbeddedPayloads(options, cancellationToken);
        return inspection.Payloads;
    }

    /// <summary>Opens a complete bounded read-only detached payload stream. The caller owns its lifetime; no external target is accessed.</summary>
    public Stream OpenEmbeddedPayload(string id, long maximumBytes, CancellationToken cancellationToken = default) =>
        OfficeOpenXmlPackagePayload.OpenDetached(_wordprocessingDocument, id, maximumBytes, cancellationToken);
}
