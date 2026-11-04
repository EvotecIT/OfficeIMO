using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    private readonly System.Collections.Generic.Dictionary<string, OfficeFontFace> _installedFamilySources = new();
    private readonly System.Collections.Generic.Dictionary<string, OfficeFontFace> _installedSourceFaces = new();
    // Resolve the installed face through the same family list as raster fallback, then
    // register its standalone program through the ordinary bounded/provider-aware loader.
    internal bool TryAddInstalledFamily(string familyName, OfficeFontFaceDescriptor requested,
        int maximumDecodedBytes, CancellationToken cancellationToken, out int decodedBytes, out string? error,
        long maximumSourceBytes = long.MaxValue) {
        decodedBytes = 0;
        error = null;
        cancellationToken.ThrowIfCancellationRequested();
        if (OfficeSystemFontFamilyAliases.IsMath(familyName)) return TryAddInstalledMathematicalFamily(
            requested, maximumDecodedBytes, maximumSourceBytes, cancellationToken, out decodedBytes, out error);
        OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadFontFamilyForText(familyName, requested, null, out _, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (font == null) {
            foreach (string family in OfficeSystemFontFamilyAliases.Expand(familyName)) {
                if (TryAddInstalledCffFamily(family, requested, OfficeTrueTypeFont.CandidateFamilyPaths(family, cancellationToken),
                        maximumDecodedBytes, maximumSourceBytes, cancellationToken, out decodedBytes, out error)) {
                    if (!string.Equals(family, familyName, StringComparison.OrdinalIgnoreCase)) AddAlias(familyName, family);
                    return true;
                }
                if (error != null) return false;
            }
            return false;
        }
        bool isGenericAlias = OfficeSystemFontFamilyAliases.IsSystemUi(familyName);
        bool nativeSelection = FontProgramProvider == null && FontVariationResolver == null;
        string sourceKey = font.Fingerprint + "#" + requested.Weight + "#" + requested.Slant;
        if (nativeSelection && _installedFamilySources.TryGetValue(sourceKey, out OfficeFontFace? registered)) {
            string aliasResource = registered.Descriptor == OfficeFontFaceDescriptor.Regular
                ? familyName : CreateResourceFamilyName(familyName, registered.Descriptor, registered.UnicodeRanges);
            OfficeFontFace alias = registered.CreateAlias(familyName, aliasResource);
            alias.InstalledSubstituteFamily = isGenericAlias || font.MatchesFamilyName(familyName)
                ? null : font.DisplayName ?? "installed substitute";
            int index = _faces.FindLastIndex(face => face.FamilyName == familyName && face.ResourceFamilyName == aliasResource);
            if (index >= 0) FacesToChange[index] = alias;
            else FacesToChange.Add(alias);
            return true;
        }
        bool shareData = nativeSelection
            && _installedSourceFaces.TryGetValue(font.Fingerprint, out _);
        if (!shareData && font.SourceByteCount > maximumSourceBytes) {
            error = "Installed font source exceeds the per-resource byte limit.";
            return false;
        }
        if (!shareData && font.SourceByteCount > maximumDecodedBytes) {
            error = "Installed font source exceeds the byte limit.";
            return false;
        }
        var templateKey = new InstalledFaceKey(font.Fingerprint, requested.Weight, requested.Slant, shareData);
        if (nativeSelection && InstalledFaceTemplates.TryGetValue(templateKey, out InstalledFaceTemplate? template)
            && template.DecodedBytes <= maximumDecodedBytes) {
            OfficeFontFace shared = RegisterInstalledTemplate(familyName, template);
            shared.InstalledSubstituteFamily = isGenericAlias || font.MatchesFamilyName(familyName)
                ? null : font.DisplayName ?? "installed substitute";
            decodedBytes = template.DecodedBytes;
            _installedFamilySources[sourceKey] = shared;
            _installedSourceFaces[font.Fingerprint] = shared;
            return true;
        }
        byte[] data;
        try {
            data = shareData ? _installedSourceFaces[font.Fingerprint].DataSnapshot
                : font.CopyStandaloneData(maximumDecodedBytes / 2);
        } catch (Exception exception) when (exception is NotSupportedException || exception is OverflowException) {
            error = exception.Message;
            return false;
        }
        cancellationToken.ThrowIfCancellationRequested();
        // Static faces retain the style they actually supply; variable wght faces select
        // the requested weight through the existing descriptor-aware loading contract.
        OfficeFontFaceDescriptor descriptor = font.FaceDescriptor;
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
        bool variable = reader != null && reader.TryGetTable("fvar", out _, out _);
        if (variable) descriptor = new OfficeFontFaceDescriptor(requested.Weight, descriptor.StretchPercent, descriptor.Slant);
        string resourceFamily = descriptor == OfficeFontFaceDescriptor.Regular
            ? familyName : CreateResourceFamilyName(familyName, descriptor, OfficeFontUnicodeRangeSet.All);
        if (!TryAddCore(familyName, data, descriptor.ToStyle(), descriptor, OfficeFontUnicodeRangeSet.All,
                resourceFamily, maximumDecodedBytes, out decodedBytes, out error,
                applyDescriptorWeight: variable, useOwnedDataSnapshot: shareData)) return false;
        OfficeFontFace registeredFace = _faces.FindLast(face =>
            face.FamilyName == familyName && face.ResourceFamilyName == resourceFamily)!;
        registeredFace.InstalledSubstituteFamily = isGenericAlias || font.MatchesFamilyName(familyName)
            ? null : font.DisplayName ?? "installed substitute";
        if (nativeSelection) {
            _installedFamilySources[sourceKey] = registeredFace;
            _installedSourceFaces[font.Fingerprint] = registeredFace;
            if (!variable) RememberInstalledTemplate(templateKey, registeredFace, decodedBytes);
        }
        return true;
    }

    private OfficeFontFace RegisterInstalledTemplate(string familyName, InstalledFaceTemplate template) {
        OfficeFontFace source = template.Face;
        string resourceFamily = source.Descriptor == OfficeFontFaceDescriptor.Regular
            ? familyName : CreateResourceFamilyName(familyName, source.Descriptor, OfficeFontUnicodeRangeSet.All);
        OfficeFontFace face = source.CreateAlias(familyName, resourceFamily);
        int index = _faces.FindLastIndex(existing => existing.Style == face.Style
            && string.Equals(existing.FamilyName, familyName, StringComparison.OrdinalIgnoreCase)
            && string.Equals(existing.ResourceFamilyName, resourceFamily, StringComparison.OrdinalIgnoreCase));
        if (index >= 0) FacesToChange[index] = face;
        else FacesToChange.Add(face);
        return face;
    }

    private static void RememberInstalledTemplate(InstalledFaceKey key, OfficeFontFace face, int decodedBytes) {
        long retained = face.DataSnapshot.Length;
        if (Interlocked.Add(ref _installedTemplateBytes, retained) > MaximumInstalledTemplateBytes
            || !InstalledFaceTemplates.TryAdd(key, new InstalledFaceTemplate(face, decodedBytes))) {
            Interlocked.Add(ref _installedTemplateBytes, -retained);
        }
    }

    // Every render would otherwise extract, decode and parse the same installed files again,
    // and the new arrays would also miss the PDF font caches, which key on them. Variable
    // programs are excluded because they fill a per-program cache while measuring.
    private const long MaximumInstalledTemplateBytes = 128L * 1024 * 1024;
    private static long _installedTemplateBytes;
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<InstalledFaceKey, InstalledFaceTemplate> InstalledFaceTemplates =
        new System.Collections.Concurrent.ConcurrentDictionary<InstalledFaceKey, InstalledFaceTemplate>();

    private readonly struct InstalledFaceKey : IEquatable<InstalledFaceKey> {
        private readonly string _fingerprint;
        private readonly int _weight;
        private readonly OfficeFontSlant _slant;
        private readonly bool _sharedData;

        internal InstalledFaceKey(string fingerprint, int weight, OfficeFontSlant slant, bool sharedData) {
            _fingerprint = fingerprint;
            _weight = weight;
            _slant = slant;
            _sharedData = sharedData;
        }

        public bool Equals(InstalledFaceKey other) => _weight == other._weight && _slant == other._slant
            && _sharedData == other._sharedData && string.Equals(_fingerprint, other._fingerprint, StringComparison.Ordinal);
        public override bool Equals(object? obj) => obj is InstalledFaceKey other && Equals(other);
        public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(_fingerprint) ^ (_weight * 397) ^ ((int)_slant << 20);
    }

    private sealed class InstalledFaceTemplate {
        internal InstalledFaceTemplate(OfficeFontFace face, int decodedBytes) {
            Face = face;
            DecodedBytes = decodedBytes;
        }

        internal OfficeFontFace Face { get; }
        internal int DecodedBytes { get; }
    }
}
