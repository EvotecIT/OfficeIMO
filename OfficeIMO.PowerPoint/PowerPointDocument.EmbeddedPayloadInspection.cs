using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    /// <summary>Inventories the VBA part and signature relationship without reading its bytes or decoding module names. Payload length remains unknown.</summary>
    public OfficeVbaProjectInfo? InspectVbaProjectMetadata() {
        ThrowIfDisposed();
        if (TryReadPackageVbaProject(maximumBytes: null, inspectSignature: true, out var project, out _)) return project?.Describe();
        return SdkPresentationPart.VbaProjectPart is { } part ? OfficeOpenXmlPackagePayload.CreateVbaDescriptor(part) : null;
    }

    /// <summary>Inventories inert native payloads and creates a bounded cumulative read session. Keep this presentation open for new reads; returned streams are detached and caller-owned.</summary>
    public OfficeEmbeddedPayloadInspection InspectEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions? options = null, CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        if (GetPackagePayloads(cancellationToken) is { } payloads) {
            return new OfficeEmbeddedPayloadInspection(payloads.Select(CreatePackagePayloadSource), options, cancellationToken);
        }
        return OfficeOpenXmlPackagePayload.Inspect(OpenXmlDocument, options, cancellationToken);
    }

    /// <summary>Inventories payloads within explicit count/member/aggregate/work ceilings without scanning unrequested bytes unless hashing is requested.</summary>
    public IReadOnlyList<OfficeEmbeddedPayloadInfo> GetEmbeddedPayloads(OfficeEmbeddedPayloadInspectionOptions options, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(options);
        using OfficeEmbeddedPayloadInspection inspection = InspectEmbeddedPayloads(options, cancellationToken);
        return inspection.Payloads;
    }

    /// <summary>Opens a complete bounded read-only detached payload stream. The caller owns its lifetime; no external target is accessed.</summary>
    public Stream OpenEmbeddedPayload(string id, long maximumBytes, CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        ArgumentOutOfRangeException.ThrowIfNegative(maximumBytes);
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentException.ThrowIfNullOrWhiteSpace(id);
        if (GetPackagePayloads(cancellationToken) is { } payloads) {
            OfficeEmbeddedPayloadInfo info = payloads.FirstOrDefault(candidate => string.Equals(candidate.Id, id, StringComparison.Ordinal))
                ?? throw new KeyNotFoundException($"Embedded package payload '{id}' was not found.");
            using var inspection = new OfficeEmbeddedPayloadInspection(new[] { CreatePackagePayloadSource(info) },
                new OfficeEmbeddedPayloadInspectionOptions {
                    MaxPayloads = 1, MaxPayloadBytes = maximumBytes, MaxTotalPayloadBytes = maximumBytes, MaxReadBytes = maximumBytes
                }, cancellationToken);
            return inspection.OpenRead(0, cancellationToken);
        }
        return OfficeOpenXmlPackagePayload.OpenDetached(OpenXmlDocument, id, maximumBytes, cancellationToken);
    }
}
