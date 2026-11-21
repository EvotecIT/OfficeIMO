using System.Security.Cryptography;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    /// <summary>
    /// Adds a document-supplied face, such as CSS <c>@font-face</c> data, reusing the parsed static
    /// program of an earlier face with the same bytes. Documents decode those bytes into a new
    /// buffer on every conversion, so this looks the program up by content rather than by array.
    /// Each operation still passes its own budget: a remembered face is used only when its
    /// ordinary decoded-byte charge fits, and every other request takes the full loading path
    /// with its limit checks, errors and messages.
    /// </summary>
    internal bool TryAddContentBounded(string? familyName, byte[]? data, OfficeFontFaceDescriptor descriptor,
        OfficeFontUnicodeRangeSet? unicodeRanges, int maximumDecodedBytes, out int decodedBytes, out string? error) {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumDecodedBytes);
        OfficeFontUnicodeRangeSet ranges = unicodeRanges ?? OfficeFontUnicodeRangeSet.All;
        if (FontProgramProvider != null || FontVariationResolver != null
            || string.IsNullOrWhiteSpace(familyName) || data == null || data.Length == 0
            || data.LongLength > MaximumContentTemplateBytes) {
            return TryAddBounded(familyName, data, descriptor, ranges, maximumDecodedBytes,
                out decodedBytes, out error, applyDescriptorWeight: true);
        }

        string normalizedFamily = familyName!.Trim();
        string resourceFamily = descriptor == OfficeFontFaceDescriptor.Regular && ranges.IsAll
            ? normalizedFamily
            : CreateResourceFamilyName(normalizedFamily, descriptor, ranges);
        ContentKey key = ContentKey.Create(data);
        if (ContentTemplates.TryGet(key, out FaceTemplate template) && template.DecodedBytes <= maximumDecodedBytes) {
            OfficeFontFace source = template.Face;
            // Only static programs are remembered, so the descriptor never selects an instance.
            ReplaceOrAddFace(new OfficeFontFace(normalizedFamily, resourceFamily, source.DataSnapshot,
                descriptor.ToStyle(), descriptor, ranges, source.ParsedFont, source.ContainerFormat,
                source.CanEmbedAsStaticPdfFont, useDataSnapshot: true));
            decodedBytes = template.DecodedBytes;
            error = null;
            return true;
        }

        if (!TryAddBounded(familyName, data, descriptor, ranges, maximumDecodedBytes,
                out decodedBytes, out error, applyDescriptorWeight: true)) return false;
        OfficeFontFace? added = _faces.FindLast(candidate => candidate.FamilyName == normalizedFamily
            && candidate.ResourceFamilyName == resourceFamily && candidate.Descriptor == descriptor);
        if (added != null && IsStaticNativeTemplateCandidate(added)) {
            ContentTemplates.Add(key, new FaceTemplate(added.Clone(), decodedBytes));
        }
        return true;
    }

    // Matches the bounded loader: the last face with the same style, family and resource family wins.
    private void ReplaceOrAddFace(OfficeFontFace face) {
        for (int index = _faces.Count - 1; index >= 0; index--) {
            OfficeFontFace existing = _faces[index];
            if (existing.Style == face.Style
                && string.Equals(existing.FamilyName, face.FamilyName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(existing.ResourceFamilyName, face.ResourceFamilyName, StringComparison.OrdinalIgnoreCase)) {
                FacesToChange[index] = face;
                return;
            }
        }
        FacesToChange.Add(face);
    }

    // The same programs the snapshot templates share: static built-in faces without tracking.
    private static bool IsStaticNativeTemplateCandidate(OfficeFontFace face) {
        bool staticNative = face.CanEmbedAsStaticPdfFont &&
            (face.ParsedFont is OfficeTrueTypeFont trueType && !trueType.IsVariableFont
                || face.ParsedFont is OfficeOpenTypeCffFont cff && !cff.IsVariable);
        if (!staticNative) return false;
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(face.DataSnapshot);
        return reader != null && !reader.TryGetTable("trak", out _, out _);
    }

    private const long MaximumContentTemplateBytes = 64L * 1024 * 1024;
    // Least recently used faces go first, and the older half is dropped under high memory load.
    private static readonly OfficeSizedLruCache<ContentKey, FaceTemplate> ContentTemplates =
        new(MaximumContentTemplateBytes, static (_, template) => template.DecodedBytes);

    private readonly struct ContentKey : IEquatable<ContentKey> {
        private readonly string _digest;
        private readonly int _length;

        private ContentKey(string digest, int length) {
            _digest = digest;
            _length = length;
        }

        internal static ContentKey Create(byte[] data) {
            using HashAlgorithm hash = SHA256.Create();
            return new ContentKey(Convert.ToBase64String(hash.ComputeHash(data)), data.Length);
        }

        public bool Equals(ContentKey other) =>
            _length == other._length && string.Equals(_digest, other._digest, StringComparison.Ordinal);
        public override bool Equals(object? obj) => obj is ContentKey other && Equals(other);
        public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(_digest) ^ _length;
    }
}
