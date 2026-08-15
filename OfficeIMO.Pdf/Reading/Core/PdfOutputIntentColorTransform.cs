using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Lazily resolves the first catalog destination output profile and applies its shared ICC
/// soft-proof transform. Profile bytes are not decoded until rendering or diagnostics needs them.
/// </summary>
internal sealed class PdfOutputIntentColorTransform {
    private readonly PdfStream? _profileStream;
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly int _maxDecodedStreamBytes;
    private readonly Lazy<OfficeIccColorProfile?> _profile;

    private PdfOutputIntentColorTransform(
        PdfStream? profileStream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        string subject) {
        _profileStream = profileStream;
        _objects = objects;
        _maxDecodedStreamBytes = maxDecodedStreamBytes;
        Subject = subject;
        _profile = new Lazy<OfficeIccColorProfile?>(
            ReadProfile,
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
    }

    internal string Subject { get; }

    internal bool IsSupported => TryGetProfile(out _);

    internal OfficeColor Apply(OfficeColor color, OfficeIccRenderingIntent renderingIntent) =>
        TryGetProfile(out OfficeIccColorProfile? profile) &&
        profile != null &&
        profile.TrySoftProof(color, renderingIntent, out OfficeColor proofed)
            ? proofed
            : color;

    internal OfficeColor Apply(
        PdfPageColorSpace colorSpace,
        IReadOnlyList<double> components,
        OfficeColor fallbackColor,
        OfficeIccRenderingIntent renderingIntent) {
        if (TryGetProfile(out OfficeIccColorProfile? profile) &&
            profile != null &&
            IsMatchingDeviceColorSpace(colorSpace, profile.ComponentCount) &&
            profile.TryConvert(components, renderingIntent, out OfficeColor converted)) {
            return converted;
        }
        return Apply(fallbackColor, renderingIntent);
    }

    internal static PdfOutputIntentColorTransform? TryCreate(
        PdfDictionary? catalog,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes) {
        if (catalog == null ||
            !catalog.Items.TryGetValue("OutputIntents", out PdfObject? outputIntentsObject)) return null;
        if (!TryResolve(objects, outputIntentsObject, out PdfObject? resolvedOutputIntents)) {
            return new PdfOutputIntentColorTransform(null, objects, maxDecodedStreamBytes, "catalog");
        }
        if (resolvedOutputIntents is PdfNull) return null;
        if (resolvedOutputIntents is PdfArray { Items.Count: 0 }) return null;
        if (resolvedOutputIntents is not PdfArray outputIntents) {
            return new PdfOutputIntentColorTransform(null, objects, maxDecodedStreamBytes, "catalog");
        }

        string? malformedSubject = null;
        for (int index = 0; index < outputIntents.Items.Count; index++) {
            PdfObject item = outputIntents.Items[index];
            string subject = item is PdfReference reference
                ? reference.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : "output-intent[" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
            if (!TryResolve(objects, item, out PdfObject? resolvedItem) || resolvedItem is not PdfDictionary outputIntent) {
                if (resolvedItem is not PdfNull) malformedSubject ??= subject;
                continue;
            }
            if (!outputIntent.Items.TryGetValue("DestOutputProfile", out PdfObject? profileObject)) continue;
            if (!TryResolve(objects, profileObject, out PdfObject? resolvedProfile)) {
                malformedSubject ??= subject;
                continue;
            }
            if (resolvedProfile is PdfNull) continue;
            if (resolvedProfile is PdfStream profileStream) {
                return new PdfOutputIntentColorTransform(profileStream, objects, maxDecodedStreamBytes, subject);
            }
            malformedSubject ??= subject;
        }

        return malformedSubject == null
            ? null
            : new PdfOutputIntentColorTransform(null, objects, maxDecodedStreamBytes, malformedSubject);
    }

    private bool TryGetProfile(out OfficeIccColorProfile? profile) {
        profile = _profile.Value;
        return profile != null;
    }

    private OfficeIccColorProfile? ReadProfile() {
        if (_profileStream == null ||
            !PdfIccProfileCache.TryRead(_profileStream, _objects, _maxDecodedStreamBytes, out OfficeIccColorProfile? profile) ||
            profile == null ||
            profile.ComponentCount is not (3 or 4) ||
            !HasCompatibleDeclaredComponentCount(profile.ComponentCount) ||
            !profile.HasOutputTransform) {
            return null;
        }

        return profile;
    }

    private bool HasCompatibleDeclaredComponentCount(int profileComponentCount) {
        if (_profileStream == null ||
            !_profileStream.Dictionary.Items.TryGetValue("N", out PdfObject? componentCountObject)) return true;
        if (!TryResolve(_objects, componentCountObject, out PdfObject? resolved)) return false;
        return resolved is PdfNull ||
               (resolved is PdfNumber componentCount && componentCount.Value == profileComponentCount);
    }

    private static bool IsMatchingDeviceColorSpace(PdfPageColorSpace colorSpace, int componentCount) =>
        (colorSpace.IsNativeDeviceRgb && componentCount == 3) ||
        (colorSpace.IsNativeDeviceCmyk && componentCount == 4);

    private static bool TryResolve(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        out PdfObject? resolved) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        resolved = value;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) {
                resolved = null;
                return false;
            }
            resolved = indirect.Value;
        }
        return true;
    }
}
