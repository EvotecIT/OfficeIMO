using OfficeIMO.Core.Internal;

namespace OfficeIMO.OpenDocument;

/// <summary>Preserves native ODF admission identity without treating corrupt data as a quota.</summary>
internal static class OdfAdmission {
    internal static OfficeInspectionResourceLimitException Limit(string name, long maximum, Exception? inner = null) =>
        new OfficeInspectionResourceLimitException("opendocument", name, maximum, inner);

    internal static byte[] ReadSnapshot(Stream stream, long maximum) {
        try { return OfficeStreamReader.ReadAllBytes(stream, maximum); }
        catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
            throw Limit(nameof(OdfLoadOptions.MaxPackageBytes), maximum, exception);
        }
    }
}
