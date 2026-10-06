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
        OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadFontFamily(familyName, requested.ToStyle(), out OfficeFontStyle resolvedStyle);
        cancellationToken.ThrowIfCancellationRequested();
        if (font == null) return false;
        bool nativeSelection = FontProgramProvider == null && FontVariationResolver == null;
        string sourceKey = font.Fingerprint + "#" + requested.Weight + "#" + requested.Slant;
        if (nativeSelection && _installedFamilySources.TryGetValue(sourceKey, out OfficeFontFace? registered)) {
            string aliasResource = registered.Descriptor == OfficeFontFaceDescriptor.Regular
                ? familyName : CreateResourceFamilyName(familyName, registered.Descriptor, registered.UnicodeRanges);
            OfficeFontFace alias = registered.CreateAlias(familyName, aliasResource);
            int index = _faces.FindLastIndex(face => face.FamilyName == familyName && face.ResourceFamilyName == aliasResource);
            if (index >= 0) _faces[index] = alias;
            else _faces.Add(alias);
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
        OfficeFontFaceDescriptor descriptor = OfficeFontFaceDescriptor.FromStyle(resolvedStyle);
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
        bool variable = reader != null && reader.TryGetTable("fvar", out _, out _);
        if (variable) descriptor = new OfficeFontFaceDescriptor(requested.Weight, descriptor.StretchPercent, descriptor.Slant);
        string resourceFamily = descriptor == OfficeFontFaceDescriptor.Regular
            ? familyName : CreateResourceFamilyName(familyName, descriptor, OfficeFontUnicodeRangeSet.All);
        if (!TryAddCore(familyName, data, descriptor.ToStyle(), descriptor, OfficeFontUnicodeRangeSet.All,
                resourceFamily, maximumDecodedBytes, out decodedBytes, out error,
                applyDescriptorWeight: variable, useOwnedDataSnapshot: shareData)) return false;
        if (nativeSelection) {
            OfficeFontFace registeredFace = _faces.FindLast(face =>
                face.FamilyName == familyName && face.ResourceFamilyName == resourceFamily)!;
            _installedFamilySources[sourceKey] = registeredFace;
            _installedSourceFaces[font.Fingerprint] = registeredFace;
        }
        return true;
    }
}
