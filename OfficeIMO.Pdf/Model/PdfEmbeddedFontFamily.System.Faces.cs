using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfEmbeddedFontFamily {
    private const int MaxSystemFaceCollections = 16;
    private const int MaxSelectedSystemFaces = 64;
    private static readonly System.Collections.Generic.Dictionary<string, System.Lazy<OfficeFontFaceCollection?>> SystemFaceCollections =
        new(System.StringComparer.Ordinal);
    private static readonly object SystemFaceCollectionsLock = new();
    private static readonly System.Collections.Generic.Dictionary<OfficeFontFace, PdfEmbeddedFontFamily> SelectedSystemFaces =
        new();
    private static readonly object SelectedSystemFacesLock = new();

    internal static bool TryResolveSystemFace(
        string familyName,
        OfficeFontFaceDescriptor descriptor,
        string text,
        out PdfEmbeddedFontFamily? selected) {
        selected = null;
        if (string.IsNullOrWhiteSpace(familyName) || string.IsNullOrEmpty(text)) return false;
        OfficeFontFaceCollection? collection = GetSystemFaceCollection(familyName);
        return collection != null
            && TrySelectSystemFace(collection, familyName, descriptor, text, out selected);
    }

    internal static bool TryMeasureSystemFaceVerticalMetrics(
        string familyName,
        OfficeFontFaceDescriptor descriptor,
        string text,
        double fontSize,
        out double height,
        out double baselineOffset) {
        height = 0D;
        baselineOffset = 0D;
        if (string.IsNullOrWhiteSpace(familyName) || string.IsNullOrEmpty(text)
            || fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) return false;
        OfficeFontFaceCollection? collection = GetSystemFaceCollection(familyName);
        if (collection == null
            || !collection.TryResolveFaceForText(text, familyName, descriptor, out OfficeFontFace? face)
            || face?.Program is not IOfficeFontBaselineMetrics metrics) return false;
        height = face.Program.LineHeight(fontSize);
        baselineOffset = metrics.BaselineOffset(fontSize);
        return true;
    }

    private static OfficeFontFaceCollection? GetSystemFaceCollection(string familyName) {
        string normalizedFamily = NormalizeFamilyKey(familyName);
        System.Lazy<OfficeFontFaceCollection?> collection;
        lock (SystemFaceCollectionsLock) {
            if (!SystemFaceCollections.TryGetValue(normalizedFamily, out collection!)) {
                collection = new System.Lazy<OfficeFontFaceCollection?>(
                    () => LoadSystemFaceCollection(familyName.Trim()),
                    System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
                if (SystemFaceCollections.Count < MaxSystemFaceCollections) {
                    SystemFaceCollections.Add(normalizedFamily, collection);
                }
            }
        }

        return collection.Value;
    }

    internal static bool TryResolveSystemFaceFromFiles(
        string familyName,
        System.Collections.Generic.IEnumerable<string> fontFiles,
        OfficeFontFaceDescriptor descriptor,
        string text,
        out PdfEmbeddedFontFamily? selected) {
        selected = null;
        if (string.IsNullOrWhiteSpace(familyName) || string.IsNullOrEmpty(text)) return false;
        OfficeFontFaceCollection? collection = LoadSystemFaceCollection(familyName, fontFiles);
        return collection != null
            && TrySelectSystemFace(collection, familyName, descriptor, text, out selected);
    }

    internal static bool TrySelectSystemFace(
        OfficeFontFaceCollection collection,
        string familyName,
        OfficeFontFaceDescriptor descriptor,
        string text,
        out PdfEmbeddedFontFamily? selected) {
        selected = null;
        if (!collection.TryResolveFaceForText(text, familyName, descriptor, out OfficeFontFace? face)) return false;
        if (face == null) return false;
        lock (SelectedSystemFacesLock) {
            if (SelectedSystemFaces.TryGetValue(face, out selected)) return true;
        }
        byte[] data = face.Data;
        string faceName = TryReadTrueTypeNameMetadata(data, out TrueTypeNameMetadata? metadata)
            ? metadata?.FullName ?? string.Empty
            : string.Empty;
        if (string.IsNullOrWhiteSpace(faceName)
            || string.Equals(faceName, familyName, System.StringComparison.OrdinalIgnoreCase)) {
            faceName = familyName + " Weight " + face.Descriptor.Weight.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " " + face.Descriptor.Slant;
        }
#if NET6_0_OR_GREATER
        string programId = System.Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(data));
#else
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        string programId = System.BitConverter.ToString(sha256.ComputeHash(data)).Replace("-", string.Empty);
#endif
        string boundedFaceName = faceName.Length <= 32 ? faceName : faceName.Substring(0, 32);
        selected = new PdfEmbeddedFontFamily("OfficeIMO-SystemFace-" + boundedFaceName + "-" + programId, data);
        lock (SelectedSystemFacesLock) {
            if (SelectedSystemFaces.Count < MaxSelectedSystemFaces) SelectedSystemFaces[face] = selected;
        }
        return true;
    }

    private static OfficeFontFaceCollection? LoadSystemFaceCollection(string familyName) =>
        LoadSystemFaceCollection(familyName, SystemFontIndex.Value.Find(NormalizeFamilyKey(familyName)));

    private static OfficeFontFaceCollection? LoadSystemFaceCollection(
        string familyName,
        System.Collections.Generic.IEnumerable<string> fontFiles) {
        string normalizedFamily = NormalizeFamilyKey(familyName);
        string[] acceptedPrefixes = BuildAcceptedFileNamePrefixes(normalizedFamily);
        var collection = new OfficeFontFaceCollection();
        int inspectedFiles = 0;
        foreach (string fontFile in fontFiles) {
            if (inspectedFiles++ >= MaxSystemFontFilesToInspect) break;
            if (!TryReadSystemFontFaces(fontFile, normalizedFamily, acceptedPrefixes,
                    out System.Collections.Generic.List<SystemFontFaceCandidate>? candidates)
                || candidates == null) continue;
            foreach (SystemFontFaceCandidate candidate in candidates) {
                OfficeFontFaceDescriptor descriptor = ReadSystemFaceDescriptor(candidate);
                collection.TryAdd(familyName, candidate.Data, descriptor);
            }
        }
        return collection.Faces.Count == 0 ? null : collection;
    }

    private static OfficeFontFaceDescriptor ReadSystemFaceDescriptor(SystemFontFaceCandidate candidate) {
        int weight = candidate.Kind is FontFaceKind.Bold or FontFaceKind.BoldItalic ? 700 : 400;
        double stretch = 100D;
        try {
            System.Collections.Generic.Dictionary<string, FontTableRecord> tables = ReadFontTableDirectory(candidate.Data);
            if (tables.TryGetValue("OS/2", out FontTableRecord os2) && os2.Length >= 8) {
                int declaredWeight = ReadUInt16(candidate.Data, os2.Offset + 4);
                if (declaredWeight >= 1 && declaredWeight <= 1000) weight = declaredWeight;
                int widthClass = ReadUInt16(candidate.Data, os2.Offset + 6);
                double[] stretches = { 0D, 50D, 62.5D, 75D, 87.5D, 100D, 112.5D, 125D, 150D, 200D };
                if (widthClass >= 1 && widthClass <= 9) stretch = stretches[widthClass];
            }
        } catch (System.NotSupportedException) {
            // The validated font program may omit optional style tables.
        }

        OfficeFontSlant slant = candidate.Kind is FontFaceKind.Italic or FontFaceKind.BoldItalic
            ? OfficeFontSlant.Italic
            : OfficeFontSlant.Normal;
        if (TryReadTrueTypeNameMetadata(candidate.Data, out TrueTypeNameMetadata? metadata)
            && metadata != null) {
            // Some collections expose "Regular" as their typographic subfamily for
            // every face. Use all face labels so an italic cannot replace its upright
            // sibling under the same numeric weight in the selection collection.
            string style = NormalizeFamilyKey(
                (metadata.TypographicSubfamilyName ?? string.Empty) + " "
                + (metadata.SubfamilyName ?? string.Empty) + " "
                + (metadata.FullName ?? string.Empty));
            if (style.Contains("oblique")) slant = OfficeFontSlant.Oblique;
            else if (style.Contains("italic")) slant = OfficeFontSlant.Italic;
        }
        return new OfficeFontFaceDescriptor(weight, stretch, slant);
    }
}
