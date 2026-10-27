using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    /// <summary>Gets inert embedded package, OLE and ActiveX descriptors from the loaded presentation.</summary>
    public IReadOnlyList<OfficeEmbeddedPayloadInfo> GetEmbeddedPayloads(bool includeSha256 = false) {
        ThrowIfDisposed();
        return OfficeOpenXmlPackagePayload.FindEmbeddedPayloads(OpenXmlDocument)
            .Select(handle => OfficeOpenXmlPackagePayload.CreateInfo(handle, includeSha256)).ToArray();
    }

    /// <summary>Inspects attached VBA metadata without executing code, within an explicit decoded-size limit.</summary>
    public OfficeVbaProjectInfo? InspectVbaProject(bool includeSha256 = false, long maxBytes = OfficeVbaProjectInfo.DefaultMaximumProjectBytes) {
        ThrowIfDisposed();
        if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
        VbaProjectPart? part = _presentationPart.VbaProjectPart;
        return part == null ? null : OfficeOpenXmlPackagePayload.CreateVbaProjectInfo(part, includeSha256, maxBytes);
    }
}
