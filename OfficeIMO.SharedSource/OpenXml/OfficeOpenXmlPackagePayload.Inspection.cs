using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlPackagePayload {
    internal static OfficeVbaProjectInfo CreateVbaDescriptor(VbaProjectPart part) => new OfficeVbaProjectInfo(
        part.Uri.ToString(), part.ContentType, -1, null, Array.Empty<string>(),
        part.Parts.Any(pair => pair.OpenXmlPart.RelationshipType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0
            || pair.OpenXmlPart.ContentType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0),
        moduleInspectionRequested: false);

    internal static OfficeEmbeddedPayloadInspection Inspect(OpenXmlPackage package,
        OfficeEmbeddedPayloadInspectionOptions? options, CancellationToken cancellationToken, Action<Action>? executeRead = null) =>
        new OfficeEmbeddedPayloadInspection(FindEmbeddedPayloads(package).Select(handle => CreateSource(handle, executeRead)), options, cancellationToken);

    internal static Stream OpenDetached(OpenXmlPackage package, string id, long maximumBytes, CancellationToken token) {
        if (maximumBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        token.ThrowIfCancellationRequested();
        OfficeOpenXmlPayloadHandle handle = FindEmbeddedPayload(package, id);
        using var inspection = new OfficeEmbeddedPayloadInspection(new[] { CreateSource(handle) },
            new OfficeEmbeddedPayloadInspectionOptions {
                MaxPayloads = 1, MaxPayloadBytes = maximumBytes, MaxTotalPayloadBytes = maximumBytes, MaxReadBytes = maximumBytes
            }, token);
        return inspection.OpenRead(0, token);
    }

    private static OfficeEmbeddedPayloadSource CreateSource(OfficeOpenXmlPayloadHandle handle, Action<Action>? executeRead = null) {
        string uri = handle.Part.Uri.ToString();
        // The SDK does not expose an unopened package part's decoded length. Keep it
        // explicitly unknown instead of opening and scanning an unrequested payload.
        var info = new OfficeEmbeddedPayloadInfo(handle.Id, handle.Kind, handle.OwnerPartUri,
            handle.RelationshipId, uri, handle.Part.ContentType,
            Path.GetFileName(uri.Replace('/', Path.DirectorySeparatorChar)), -1, sha256: null);
        return new OfficeEmbeddedPayloadSource(info, token => {
            token.ThrowIfCancellationRequested();
            return handle.Part.GetStream(FileMode.Open, FileAccess.Read);
        }, executeRead);
    }
}
