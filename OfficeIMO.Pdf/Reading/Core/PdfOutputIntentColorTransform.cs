using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Lazily resolves the first catalog destination output profile and applies its shared ICC
/// soft-proof transform. Profile bytes are not decoded until rendering or diagnostics needs them.
/// </summary>
internal sealed class PdfOutputIntentColorTransform {
    private readonly PdfStream? _profileStream;
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly int _maxDecodedStreamBytes;
    private readonly PdfIccProfileRetentionBudget? _retentionBudget;
    private readonly PdfReadCache<OfficeIccColorProfile?> _profile = new();

    private PdfOutputIntentColorTransform(
        PdfStream? profileStream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        string subject,
        PdfIccProfileRetentionBudget? retentionBudget) {
        _profileStream = profileStream;
        _objects = objects;
        _maxDecodedStreamBytes = maxDecodedStreamBytes;
        _retentionBudget = retentionBudget;
        Subject = subject;
    }

    internal string Subject { get; }

    internal bool IsSupported => TryGetProfile(out _);

    /// <summary>Initializes shared profile state within the render caller's cancellation scope.</summary>
    internal void Prepare(CancellationToken cancellationToken) =>
        _ = GetProfile(cancellationToken);

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
        if (TryApplyDirect(colorSpace, components, renderingIntent, out OfficeColor converted)) return converted;
        return Apply(fallbackColor, renderingIntent);
    }

    internal bool TryApplyDirect(
        PdfPageColorSpace colorSpace,
        IReadOnlyList<double> components,
        OfficeIccRenderingIntent renderingIntent,
        out OfficeColor color) {
        color = OfficeColor.Black;
        return TryGetProfile(out OfficeIccColorProfile? profile) &&
            profile != null &&
            colorSpace.TryGetOutputProfileComponents(components, profile.ComponentCount, out IReadOnlyList<double> profileComponents) &&
            profile.TryConvert(profileComponents, renderingIntent, out color);
    }

    internal bool CanApplyDirect(PdfPageColorSpace colorSpace) =>
        TryGetProfile(out OfficeIccColorProfile? profile) &&
        profile != null &&
        colorSpace.CanMapDirectlyToOutputProfile(profile.ComponentCount);

    internal bool HasUncertifiableShadingComposition(PdfPageColorSpace colorSpace) {
        if (!TryGetProfile(out OfficeIccColorProfile? profile) || profile == null) return false;
        return colorSpace.CanMapDirectlyToOutputProfile(profile.ComponentCount)
            ? profile.HasUncertifiableShadingInputTransform
            : profile.HasUncertifiableShadingSoftProofTransform;
    }

    internal static PdfOutputIntentColorTransform? TryCreate(
        PdfDictionary? catalog,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        PdfIccProfileRetentionBudget? retentionBudget = null) {
        if (catalog == null ||
            !catalog.Items.TryGetValue("OutputIntents", out PdfObject? outputIntentsObject)) return null;
        if (!TryResolve(objects, outputIntentsObject, out PdfObject? resolvedOutputIntents)) {
            return new PdfOutputIntentColorTransform(null, objects, maxDecodedStreamBytes, "catalog", retentionBudget);
        }
        if (resolvedOutputIntents is PdfNull) return null;
        if (resolvedOutputIntents is PdfArray { Items.Count: 0 }) return null;
        if (resolvedOutputIntents is not PdfArray outputIntents) {
            return new PdfOutputIntentColorTransform(null, objects, maxDecodedStreamBytes, "catalog", retentionBudget);
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
                return new PdfOutputIntentColorTransform(profileStream, objects, maxDecodedStreamBytes, subject, retentionBudget);
            }
            malformedSubject ??= subject;
        }

        return malformedSubject == null
            ? null
            : new PdfOutputIntentColorTransform(null, objects, maxDecodedStreamBytes, malformedSubject, retentionBudget);
    }

    private bool TryGetProfile(out OfficeIccColorProfile? profile) {
        profile = GetProfile(CancellationToken.None);
        return profile != null;
    }

    private OfficeIccColorProfile? GetProfile(CancellationToken cancellationToken) =>
        _profile.GetOrCreate(this, static (owner, token) => owner.ReadProfile(token), cancellationToken);

    private OfficeIccColorProfile? ReadProfile(CancellationToken cancellationToken) {
        if (_profileStream == null ||
            !PdfIccProfileCache.TryRead(
                _profileStream,
                _objects,
                _maxDecodedStreamBytes,
                _retentionBudget,
                out OfficeIccColorProfile? profile,
                cancellationToken) ||
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
