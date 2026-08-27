using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Catalog output intent metadata discovered from /OutputIntents.</summary>
    public IReadOnlyList<PdfOutputIntentInfo> OutputIntents => ReadLogicalContent(_outputIntents);

    /// <summary>True when every entry in the catalog /OutputIntents array was resolved as an output-intent dictionary.</summary>
    public bool OutputIntentsAreComplete => ReadLogicalContent(_outputIntentsAreComplete);

    private IReadOnlyList<PdfOutputIntentInfo> ExtractOutputIntents(out bool isComplete) {
        isComplete = true;
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null || !catalog.Items.TryGetValue("OutputIntents", out PdfObject? outputIntentsObject)) {
            return Array.Empty<PdfOutputIntentInfo>();
        }
        if (ResolveArray(outputIntentsObject) is not PdfArray outputIntentsArray) {
            isComplete = false;
            return Array.Empty<PdfOutputIntentInfo>();
        }
        if (outputIntentsArray.Items.Count == 0) return Array.Empty<PdfOutputIntentInfo>();

        var result = new List<PdfOutputIntentInfo>();
        for (int i = 0; i < outputIntentsArray.Items.Count; i++) {
            PdfObject item = outputIntentsArray.Items[i];
            int? objectNumber = item is PdfReference reference ? reference.ObjectNumber : null;
            PdfDictionary? outputIntent = ResolveDict(item);
            if (outputIntent is null) {
                isComplete = false;
                continue;
            }

            PdfStream? profileStream = null;
            int? profileObjectNumber = null;
            if (outputIntent.Items.TryGetValue("DestOutputProfile", out PdfObject? profileObject)) {
                if (profileObject is PdfReference profileReference) {
                    profileObjectNumber = profileReference.ObjectNumber;
                }

                profileStream = ResolveObject(profileObject) as PdfStream;
            }

            Func<PdfOutputIntentProfileMetadata?>? metadataFactory = profileStream == null
                ? null
                : () => ReadOutputIntentProfileMetadata(profileStream);
            result.Add(new PdfOutputIntentInfo(
                objectNumber,
                TryReadName(outputIntent, "S"),
                TryReadText(outputIntent, "OutputConditionIdentifier"),
                TryReadText(outputIntent, "OutputCondition"),
                TryReadText(outputIntent, "RegistryName"),
                TryReadText(outputIntent, "Info"),
                profileObjectNumber,
                profileStream is null ? null : TryReadStreamColorComponents(profileStream),
                profileStream is null ? null : TryReadStreamAlternateColorSpace(profileStream),
                profileStream is null ? null : TryReadStreamFilter(profileStream),
                profileStream != null,
                metadataFactory));
        }

        return result.Count == 0 ? Array.Empty<PdfOutputIntentInfo>() : result.AsReadOnly();
    }

    private PdfOutputIntentProfileMetadata ReadOutputIntentProfileMetadata(PdfStream profileStream) {
        byte[] profileBytes = _decodedStreamBudget.DecodeRequired(
            profileStream,
            _objects,
            _options.Limits.MaxDecodedStreamBytes);
        _outputIntentMetadataRetentionBudget.Charge(
            profileStream,
            PdfIccProfileCacheRepresentation.DecodedBytes,
            profileBytes.LongLength);
        bool hasSupportedOutputTransform =
            OfficeIccColorProfile.TryCreate(profileBytes, out OfficeIccColorProfile? profile) &&
            profile != null &&
            profile.HasOutputTransform;
        return new PdfOutputIntentProfileMetadata(
            profileBytes.Length,
            TryReadIccDeclaredSize(profileBytes),
            TryReadIccColorSpace(profileBytes),
            TryReadIccDeviceClass(profileBytes),
            TryReadIccSignature(profileBytes),
            hasSupportedOutputTransform);
    }

    private int? TryReadStreamColorComponents(PdfStream stream) {
        if (!stream.Dictionary.Items.TryGetValue("N", out PdfObject? value) ||
            ResolveObject(value) is not PdfNumber number ||
            number.Value < 0 ||
            number.Value > int.MaxValue ||
            Math.Truncate(number.Value) != number.Value) {
            return null;
        }

        return (int)number.Value;
    }

    private string? TryReadStreamAlternateColorSpace(PdfStream stream) {
        return TryReadName(stream.Dictionary, "Alternate");
    }

    private string? TryReadStreamFilter(PdfStream stream) {
        if (!stream.Dictionary.Items.TryGetValue("Filter", out PdfObject? value) ||
            !TryFormatSimpleValue(value, out string? filter)) {
            return null;
        }

        return filter;
    }

    private static int? TryReadIccDeclaredSize(byte[] profile) {
        if (profile.Length < 4) {
            return null;
        }

        return (profile[0] << 24) |
            (profile[1] << 16) |
            (profile[2] << 8) |
            profile[3];
    }

    private static string? TryReadIccColorSpace(byte[] profile) {
        if (profile.Length < 20) {
            return null;
        }

        return Encoding.ASCII.GetString(profile, 16, 4);
    }

    private static string? TryReadIccDeviceClass(byte[] profile) {
        if (profile.Length < 16) {
            return null;
        }

        return Encoding.ASCII.GetString(profile, 12, 4);
    }

    private static bool? TryReadIccSignature(byte[] profile) {
        if (profile.Length < 40) {
            return null;
        }

        return profile[36] == (byte)'a' &&
            profile[37] == (byte)'c' &&
            profile[38] == (byte)'s' &&
            profile[39] == (byte)'p';
    }
}
