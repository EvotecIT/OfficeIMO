using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Internal;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    private const string VbaProjectContentType = "application/vnd.ms-office.vbaProject";

    private readonly record struct PackageVbaProject(string PartName, bool HasSignature) {
        internal OfficeVbaProjectInfo Describe(byte[]? data = null, bool includeSha256 = false) =>
            new(PartName, VbaProjectContentType, data?.LongLength ?? -1,
                data != null && includeSha256
                    ? OfficeRuntimeCompatibility.ToHexString(OfficeRuntimeCompatibility.HashSha256(data)).ToLowerInvariant()
                    : null,
                data == null ? Array.Empty<string>() : OfficeVbaProjectInspector.GetModuleNames(data), HasSignature,
                moduleInspectionRequested: data != null);
    }

    private bool TryReadPackageVbaProject(long? maximumBytes, bool inspectSignature,
        out PackageVbaProject? project, out byte[]? bytes, bool validateForSource = false) {
        project = null;
        bytes = null;
        if (!_sdkDeferred || IsSdkLoaded || _deferredOpenSettings?.MaxCharactersInPart > 0) return false;
        SemaphoreSlim gate = PayloadReadGate;
        gate.Wait();
        try {
            return TryReadPackageVbaProjectCore(maximumBytes, inspectSignature, out project, out bytes, validateForSource);
        } finally {
            gate.Release();
        }
    }

    private bool TryReadPackageVbaProjectCore(long? maximumBytes, bool inspectSignature,
        out PackageVbaProject? project, out byte[]? bytes, bool validateForSource) {
        project = null;
        bytes = null;
        OfficeOpcPackage? loaded = LoadedPackageForCurrentBytes;
        OfficeOpcEditablePackage? editable = loaded == null ? ReadablePackage : null;
        if (loaded == null && editable == null) return false;

        string? main;
        if (loaded != null) {
            main = loaded.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
        } else {
            main = FirstRelated(editable!, null, OfficeOpcPackage.OfficeDocumentRelationshipType, out _);
        }
        if (main == null) return false;
        string? mainPart = loaded?.GetPartName(main) ?? editable?.GetPartName(main);
        if (mainPart == null) return false;

        string? projectPart = null;
        HashSet<string>? signatureOwners = null;
        void InspectRelationship(PackageGraphRelationship relationship) {
            if (relationship.Owner == mainPart && relationship.Type == VbaProjectRelationshipType) {
                projectPart ??= relationship.Part;
            }
            if (inspectSignature && relationship.Owner is { } owner && IsVbaProjectSignature(relationship.Type, relationship.ContentType)) {
                (signatureOwners ??= new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance)).Add(owner);
            }
        }

        bool visited;
        try {
            visited = loaded != null
                ? TryVisitPackageGraph(loaded, default, InspectRelationship, validateForSave: validateForSource)
                : TryVisitPackageGraph(editable!, default, InspectRelationship, validateForSave: validateForSource);
        } catch (Exception exception) when (validateForSource && (exception is InvalidDataException or IOException or System.Xml.XmlException
            or ArgumentException or NotSupportedException or InvalidOperationException or OverflowException)) {
            return false;
        }
        if (!visited) return false;
        if (projectPart == null) return true;
        string? contentType = loaded?.GetContentType(projectPart) ?? editable?.GetContentType(projectPart);
        if (contentType != VbaProjectContentType) return false;
        if (maximumBytes.HasValue) {
            if (loaded != null) OfficeStreamReader.EnsureWithinLimit(loaded.GetPartLength(projectPart), maximumBytes);
            else if (editable!.TryGetBinaryLength(projectPart, out long length)) OfficeStreamReader.EnsureWithinLimit(length, maximumBytes);
            using Stream input = loaded?.OpenPartContentStream(projectPart, default) ?? editable!.OpenPartContentStream(projectPart, default);
            bytes = OfficeStreamReader.ReadAllBytes(input, maximumBytes);
        }
        project = new PackageVbaProject(projectPart, signatureOwners?.Contains(projectPart) == true);
        return true;
    }

    private static bool IsVbaProjectSignature(string relationshipType, string contentType) =>
        relationshipType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0 ||
        contentType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0;
}
