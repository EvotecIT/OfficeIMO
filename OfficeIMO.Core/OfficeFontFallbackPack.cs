using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Immutable, fingerprinted set of deterministic font faces and fallback order that can be reused
/// across HTML, PDF, SVG, image, and document-conversion options.
/// </summary>
public sealed class OfficeFontFallbackPack {
    private readonly OfficeFontFaceCollection _fonts;

    /// <summary>Creates a validated snapshot of one portable fallback pack.</summary>
    /// <param name="id">Stable pack identifier suitable for logs and artifact manifests.</param>
    /// <param name="defaultFamilyNames">Office/CSS family list used when the source has no explicit family.</param>
    /// <param name="fonts">Registered faces, aliases, Unicode ranges, and ordered fallback families.</param>
    public OfficeFontFallbackPack(
        string id,
        string defaultFamilyNames,
        OfficeFontFaceCollection fonts) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("A font fallback pack id is required.", nameof(id));
        if (string.IsNullOrWhiteSpace(defaultFamilyNames)) {
            throw new ArgumentException("A default font family list is required.", nameof(defaultFamilyNames));
        }
        if (fonts == null) throw new ArgumentNullException(nameof(fonts));
        if (fonts.Faces.Count == 0) throw new ArgumentException("A font fallback pack must contain at least one face.", nameof(fonts));

        Id = id.Trim();
        DefaultFamilyNames = defaultFamilyNames.Trim();
        _fonts = fonts.Clone();
        Fingerprint = ComputeFingerprint(Id, DefaultFamilyNames, _fonts);
    }

    /// <summary>Stable pack identifier.</summary>
    public string Id { get; }

    /// <summary>Default Office/CSS family list for this pack.</summary>
    public string DefaultFamilyNames { get; }

    /// <summary>Lowercase SHA-256 over the pack identity, faces, program instances, ranges, bytes, and fallback order.</summary>
    public string Fingerprint { get; }

    /// <summary>Independent snapshot of all registered faces and ordered fallback families.</summary>
    public OfficeFontFaceCollection Fonts => _fonts.Clone();

    /// <summary>Creates a shared rendering profile over this pack.</summary>
    public OfficeRenderingProfile CreateRenderingProfile(
        IOfficeTextShapingProvider? textShapingProvider = null,
        string? textShapingLanguage = null,
        IOfficeRasterImageCodec? imageCodec = null,
        OfficeImageExportPolicy? policy = null) =>
        new OfficeRenderingProfile(
            Id,
            _fonts,
            textShapingProvider,
            textShapingLanguage,
            imageCodec,
            policy);

    private static string ComputeFingerprint(
        string id,
        string defaultFamilyNames,
        OfficeFontFaceCollection fonts) {
        using HashAlgorithm hash = SHA256.Create();
        Append(hash, "OfficeIMO.FontFallbackPack.v2");
        Append(hash, id);
        Append(hash, defaultFamilyNames);
        Append(hash, fonts.Faces.Count.ToString(CultureInfo.InvariantCulture));
        foreach (OfficeFontFace face in fonts.Faces) {
            Append(hash, face.FamilyName);
            Append(hash, face.ResourceFamilyName);
            Append(hash, ((int)face.Style).ToString(CultureInfo.InvariantCulture));
            Append(hash, face.UnicodeRanges.ToStableKey());
            Append(hash, ((int)face.ContainerFormat).ToString(CultureInfo.InvariantCulture));
            Append(hash, face.CanEmbedAsStaticPdfFont ? "static-pdf" : "outlined-pdf");
            Append(hash, face.Program.Fingerprint);
            Append(hash, face.DataSnapshot);
        }
        Append(hash, fonts.FallbackFamilies.Count.ToString(CultureInfo.InvariantCulture));
        foreach (string family in fonts.FallbackFamilies) Append(hash, family);
        hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return ToLowerHex(hash.Hash!);
    }

    private static void Append(HashAlgorithm hash, string value) =>
        Append(hash, Encoding.UTF8.GetBytes(value));

    private static void Append(HashAlgorithm hash, byte[] value) {
        byte[] length = {
            (byte)(value.Length >> 24),
            (byte)(value.Length >> 16),
            (byte)(value.Length >> 8),
            (byte)value.Length
        };
        hash.TransformBlock(length, 0, length.Length, length, 0);
        if (value.Length > 0) hash.TransformBlock(value, 0, value.Length, value, 0);
    }

    private static string ToLowerHex(byte[] bytes) {
        const string alphabet = "0123456789abcdef";
        var result = new char[bytes.Length * 2];
        for (int index = 0; index < bytes.Length; index++) {
            result[index * 2] = alphabet[bytes[index] >> 4];
            result[index * 2 + 1] = alphabet[bytes[index] & 0x0F];
        }
        return new string(result);
    }
}
