using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Drawing;
using SixLabors.Fonts;
using SixLabors.Fonts.Tables.AdvancedTypographic;

namespace OfficeIMO.Drawing.SixLabors;

/// <summary>
/// Managed optional font-program provider backed by SixLabors.Fonts.
/// </summary>
public sealed class OfficeSixLaborsFontProgramProvider : IOfficeFontProgramProvider {
    private readonly Func<OfficeFontProgramLoadRequest, IReadOnlyDictionary<string, float>?>? _variationResolver;

    /// <summary>Shared provider using each variable font's default axis values.</summary>
    public static OfficeSixLaborsFontProgramProvider Instance { get; } = new();

    /// <summary>Creates a provider using default variable-font axis values.</summary>
    public OfficeSixLaborsFontProgramProvider() {
    }

    /// <summary>
    /// Creates a provider whose resolver can select variable-font axes for each loaded face.
    /// Axis tags must contain exactly four printable ASCII characters.
    /// </summary>
    public OfficeSixLaborsFontProgramProvider(
        Func<OfficeFontProgramLoadRequest, IReadOnlyDictionary<string, float>?> variationResolver) {
        _variationResolver = variationResolver ?? throw new ArgumentNullException(nameof(variationResolver));
    }

    /// <inheritdoc />
    public OfficeFontProgramLoadResult? TryLoad(OfficeFontProgramLoadRequest request) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (request.ContainerFormat != OfficeFontContainerFormat.OpenType
            && request.ContainerFormat != OfficeFontContainerFormat.Woff
            && request.ContainerFormat != OfficeFontContainerFormat.Woff2) {
            return null;
        }

        byte[] data = request.Data;
        int decodedByteCount = ResolveDecodedByteCount(data, request.ContainerFormat);
        if (decodedByteCount <= 0 || decodedByteCount > request.MaximumDecodedBytes) {
            throw new InvalidDataException("Decoded font data exceeds the configured byte limit.");
        }

        var collection = new FontCollection();
        FontFamily family;
        FontDescription description;
        using (var stream = new MemoryStream(data, writable: false)) {
            family = collection.Add(stream, out description);
        }

        FontVariation[] variations = ResolveVariations(request, out string variationIdentity);
        FontStyle style = description.Style;
        Font prototype = family.CreateFont(1F, style, variations);
        FontMetrics metrics = prototype.FontMetrics;
        bool hasCff1 = metrics.TryGetTableData(Tag.Parse("CFF "), out _);
        bool hasCff2 = metrics.TryGetTableData(Tag.Parse("CFF2"), out _);
        bool isVariable = metrics.TryGetVariationAxes(out ReadOnlyMemory<global::SixLabors.Fonts.Tables.AdvancedTypographic.Variations.VariationAxis> axes)
            && axes.Length > 0;
        byte[]? staticOpenTypeData = IsSfnt(data)
            && !hasCff2
            && !isVariable
            ? data
            : null;

        var program = new OfficeSixLaborsFontProgram(
            data,
            family,
            style,
            variations,
            description.FontNameInvariantCulture,
            hasCff1 || hasCff2,
            ComputeFingerprint(data, variationIdentity));
        return new OfficeFontProgramLoadResult(program, decodedByteCount, staticOpenTypeData);
    }

    private FontVariation[] ResolveVariations(OfficeFontProgramLoadRequest request, out string identity) {
        IReadOnlyDictionary<string, float>? values = _variationResolver?.Invoke(request);
        if (values == null || values.Count == 0) {
            identity = "defaults";
            return Array.Empty<FontVariation>();
        }

        var result = new FontVariation[values.Count];
        int index = 0;
        var identityBuilder = new StringBuilder();
        foreach (KeyValuePair<string, float> value in values.OrderBy(value => value.Key, StringComparer.Ordinal)) {
            if (value.Key == null || value.Key.Length != 4) {
                throw new ArgumentException("Variable-font axis tags must contain exactly four characters.");
            }
            for (int character = 0; character < value.Key.Length; character++) {
                if (value.Key[character] < 0x20 || value.Key[character] > 0x7E) {
                    throw new ArgumentException("Variable-font axis tags must contain printable ASCII characters.");
                }
            }
            if (float.IsNaN(value.Value) || float.IsInfinity(value.Value)) {
                throw new ArgumentOutOfRangeException(nameof(value), "Variable-font axis values must be finite.");
            }
            result[index++] = new FontVariation(value.Key, value.Value);
            if (identityBuilder.Length > 0) identityBuilder.Append(';');
            identityBuilder.Append(value.Key)
                .Append('=')
                .Append(value.Value.ToString("R", CultureInfo.InvariantCulture));
        }
        identity = identityBuilder.ToString();
        return result;
    }

    private static string ComputeFingerprint(byte[] data, string variationIdentity) {
        using HashAlgorithm hash = SHA256.Create();
        byte[] identity = Encoding.UTF8.GetBytes("OfficeIMO.SixLabors.FontProgram.v1\n" + variationIdentity);
        hash.TransformBlock(identity, 0, identity.Length, identity, 0);
        hash.TransformFinalBlock(data, 0, data.Length);
        return "sha256:" + ToLowerHex(hash.Hash!);
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

    private static int ResolveDecodedByteCount(byte[] data, OfficeFontContainerFormat format) {
        if (format != OfficeFontContainerFormat.Woff && format != OfficeFontContainerFormat.Woff2) {
            return data.Length;
        }
        if (data.Length < 20) throw new InvalidDataException("The web-font header is truncated.");
        uint declared = ReadUInt32(data, 16);
        if (declared == 0U || declared > int.MaxValue) {
            throw new InvalidDataException("The web-font decoded size is invalid.");
        }
        return checked((int)declared);
    }

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24)
        | ((uint)data[offset + 1] << 16)
        | ((uint)data[offset + 2] << 8)
        | data[offset + 3];

    private static bool IsSfnt(byte[] data) =>
        data.Length >= 12
        && ((data[0] == 0 && data[1] == 1 && data[2] == 0 && data[3] == 0)
            || (data[0] == (byte)'O' && data[1] == (byte)'T'
                && data[2] == (byte)'T' && data[3] == (byte)'O'));
}
