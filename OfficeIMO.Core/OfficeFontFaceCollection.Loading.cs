using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    private bool TryAddCore(
        string? familyName,
        byte[]? data,
        OfficeFontStyle style,
        OfficeFontFaceDescriptor descriptor,
        OfficeFontUnicodeRangeSet unicodeRanges,
        string? resourceFamilyName,
        int? maximumDecodedBytes,
        out int decodedBytes,
        out string? error,
        bool applyDescriptorWeight = false,
        bool useOwnedDataSnapshot = false) {
        decodedBytes = 0;
        error = null;
        if (string.IsNullOrWhiteSpace(familyName) || data == null || data.Length == 0) {
            error = "Font data and family name are required.";
            return false;
        }

        OfficeFontContainerFormat sourceFormat = OfficeFontContainerDecoder.Detect(data);
        byte[] openTypeData = Array.Empty<byte>();
        bool decoded = useOwnedDataSnapshot || (maximumDecodedBytes.HasValue
            ? OfficeFontContainerDecoder.TryDecodeToOpenType(
                data,
                maximumDecodedBytes.Value,
                out openTypeData,
                out _,
                out error)
            : OfficeFontContainerDecoder.TryDecodeToOpenType(
                data,
                out openTypeData,
                out _,
                out error));
        if (useOwnedDataSnapshot) openTypeData = data;
        IReadOnlyDictionary<string, float>? variationValues = null;
        if (decoded && (FontVariationResolver != null || applyDescriptorWeight)) {
            try {
                variationValues = FontVariationResolver?.Invoke(new OfficeFontProgramLoadRequest(
                    familyName!.Trim(),
                    openTypeData,
                    OfficeFontFace.NormalizeStyle(style),
                    OfficeFontContainerFormat.OpenType,
                    maximumDecodedBytes ?? OfficeFontContainerDecoder.DefaultMaximumDecodedBytes));
                if (applyDescriptorWeight) {
                    variationValues = ResolveDescriptorWeightCoordinates(openTypeData, descriptor.Weight, variationValues);
                }
            } catch (Exception exception) when (!(exception is OutOfMemoryException)) {
                error = "The variable-font axis selection failed: " + exception.Message;
                return false;
            }
        }
        bool isFontCollection = decoded && HasTrueTypeCollectionSignature(openTypeData);
        if (isFontCollection && variationValues != null && variationValues.Count > 0) {
            error = "Variable-font axes cannot be selected on a font collection. Extract and register the intended face as an individual OpenType font.";
            return false;
        }
        IOfficeFontProgram? builtInProgram = null;
        bool builtInVariable = false;
        if (decoded) {
            OfficeOpenTypeCffFont? cffProgram = OfficeOpenTypeCffFont.TryLoad(openTypeData, variationValues, out string? cffError);
            if (cffProgram != null) {
                builtInProgram = cffProgram;
                builtInVariable = cffProgram.IsVariable;
            } else {
                OfficeOpenTypeReader? openTypeReader = OfficeOpenTypeReader.TryCreate(openTypeData);
                OfficeFontVariationModel variationModel;
                try {
                    variationModel = openTypeReader == null
                        ? OfficeFontVariationModel.None
                        : OfficeFontVariationModel.Create(openTypeReader, variationValues);
                } catch (Exception exception) when (!(exception is OutOfMemoryException)) {
                    error = "The variable-font configuration is invalid: " + exception.Message;
                    return false;
                }
                builtInVariable = variationModel.IsVariable;
                string? trueTypeError = null;
                builtInProgram = variationModel.IsVariable
                    ? OfficeTrueTypeFont.TryLoad(openTypeData, variationModel, out trueTypeError)
                    : OfficeTrueTypeFont.TryLoad(openTypeData);
                if (builtInProgram == null && variationModel.IsVariable && !string.IsNullOrWhiteSpace(trueTypeError)) {
                    error = trueTypeError;
                }
                if (builtInProgram == null && !string.IsNullOrWhiteSpace(cffError)) error = cffError;
                if (builtInProgram == null && isFontCollection) {
                    error = "This font collection cannot be registered directly. Extract and register the intended face as an individual OpenType font.";
                }
            }
        }
        IOfficeFontProgram? parsed = builtInProgram;
        byte[] acceptedData = openTypeData;
        bool canEmbedAsStaticPdfFont = parsed != null && !builtInVariable && !isFontCollection;
        // Configuring a provider is an explicit request to use its complete layout engine even
        // for TrueType faces the dependency-free core can decode. A provider may still decline,
        // in which case the already validated built-in program remains the fallback.
        bool providerPreferred = decoded;
        if ((parsed == null || providerPreferred) && FontProgramProvider != null) {
            int providerLimit = maximumDecodedBytes ?? OfficeFontContainerDecoder.DefaultMaximumDecodedBytes;
            IReadOnlyDictionary<string, float>? providerVariationValues =
                (builtInProgram as IOfficeVariableFontProgram)?.VariationCoordinatesForShaping
                ?? variationValues;
            // The request's container must describe the bytes handed to the provider. WOFF inputs
            // are normalized by the core before this point. Keeping the original WOFF label with
            // sfnt bytes makes a provider interpret the table directory as a web-font header.
            OfficeFontContainerFormat providerInputFormat = decoded
                ? OfficeFontContainerDecoder.Detect(openTypeData)
                : sourceFormat;
            OfficeFontProgramLoadResult? providerResult;
            try {
                providerResult = FontProgramProvider.TryLoad(new OfficeFontProgramLoadRequest(
                    familyName!.Trim(),
                    decoded ? openTypeData : data,
                    OfficeFontFace.NormalizeStyle(style),
                    providerInputFormat,
                    providerLimit,
                    providerVariationValues));
            } catch (Exception exception) when (!(exception is OutOfMemoryException)) {
                error = "The configured font-program provider failed: " + exception.Message;
                return false;
            }
            if (providerResult != null) {
                byte[]? staticData = providerResult.StaticOpenTypeDataSnapshot;
                int faceDataBytes = staticData?.Length ?? data.Length;
                long retainedBytes = (long)providerResult.DecodedByteCount + faceDataBytes;
                if (retainedBytes > providerLimit) {
                    error = "Decoded font data exceeds the configured byte limit.";
                    return false;
                }
                parsed = providerResult.Program;
                acceptedData = staticData ?? (byte[])data.Clone();
                decodedBytes = checked((int)retainedBytes);
                canEmbedAsStaticPdfFont = staticData != null && !HasTrueTypeCollectionSignature(staticData);
            } else {
                parsed = builtInProgram;
                acceptedData = openTypeData;
                canEmbedAsStaticPdfFont = builtInProgram != null && !builtInVariable && !isFontCollection;
            }
        }
        if (parsed == null) {
            if (decoded && string.IsNullOrWhiteSpace(error)) error = "Decoded font data does not contain a supported outline program.";
            return false;
        }
        if (!useOwnedDataSnapshot && decodedBytes == 0 && maximumDecodedBytes.HasValue) {
            // The face owns one independent embedding snapshot. The built-in TrueType program
            // retains the decoded sfnt buffer, while CFF retains that reader buffer plus its
            // independent shaping snapshot.
            int retainedFullBufferCount = parsed is OfficeOpenTypeCffFont ? 3 : 2;
            long retainedBytes = (long)acceptedData.Length * retainedFullBufferCount;
            if (retainedBytes > maximumDecodedBytes.Value || retainedBytes > int.MaxValue) {
                error = "Decoded font data exceeds the configured byte limit.";
                return false;
            }
            decodedBytes = (int)retainedBytes;
        }

        string normalizedFamily = familyName!.Trim();
        string normalizedResourceFamily = string.IsNullOrWhiteSpace(resourceFamilyName)
            ? normalizedFamily
            : resourceFamilyName!.Trim();
        OfficeFontUnicodeRangeSet normalizedRanges = unicodeRanges;
        OfficeFontStyle normalizedStyle = OfficeFontFace.NormalizeStyle(style);
        for (int index = _faces.Count - 1; index >= 0; index--) {
            OfficeFontFace existing = _faces[index];
            if (existing.Style == normalizedStyle
                && string.Equals(existing.FamilyName, normalizedFamily, StringComparison.OrdinalIgnoreCase)
                && string.Equals(existing.ResourceFamilyName, normalizedResourceFamily, StringComparison.OrdinalIgnoreCase)) {
                _faces[index] = new OfficeFontFace(
                    normalizedFamily,
                    normalizedResourceFamily,
                    acceptedData,
                    normalizedStyle,
                    descriptor,
                    normalizedRanges,
                    parsed,
                    sourceFormat,
                    canEmbedAsStaticPdfFont, useDataSnapshot: useOwnedDataSnapshot,
                    automaticOpticalSizing: ReferenceEquals(parsed, builtInProgram));
                if (!useOwnedDataSnapshot && decodedBytes == 0) decodedBytes = acceptedData.Length;
                return true;
            }
        }

        _faces.Add(new OfficeFontFace(
            normalizedFamily,
            normalizedResourceFamily,
            acceptedData,
            normalizedStyle,
            descriptor,
            normalizedRanges,
            parsed,
            sourceFormat,
            canEmbedAsStaticPdfFont, useDataSnapshot: useOwnedDataSnapshot,
            automaticOpticalSizing: ReferenceEquals(parsed, builtInProgram)));
        if (!useOwnedDataSnapshot && decodedBytes == 0) decodedBytes = acceptedData.Length;
        return true;
    }

    /// <summary>Uses a CSS face weight when its variable font defines wght and the caller did not select it.</summary>
    private static IReadOnlyDictionary<string, float>? ResolveDescriptorWeightCoordinates(
        byte[] openTypeData,
        int weight,
        IReadOnlyDictionary<string, float>? requested) {
        if (requested != null && requested.ContainsKey("wght")) return requested;
        OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(openTypeData);
        if (reader == null) return requested;
        OfficeFontVariationModel model = OfficeFontVariationModel.Create(reader, null);
        if (!model.DesignCoordinates.ContainsKey("wght")) return requested;
        var coordinates = new Dictionary<string, float>(StringComparer.Ordinal);
        if (requested != null) {
            foreach (KeyValuePair<string, float> coordinate in requested) coordinates.Add(coordinate.Key, coordinate.Value);
        }
        coordinates.Add("wght", weight);
        return coordinates;
    }
}
