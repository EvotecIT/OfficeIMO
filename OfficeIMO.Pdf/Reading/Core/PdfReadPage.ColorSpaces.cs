using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private const int MaxColorSpaceNesting = 8;
    private const int MaxDeviceNComponents = 32;

    private bool TryReadExtendedColorSpaceResource(
        PdfObject? value,
        int depth,
        Func<int, bool>? evaluationBudget,
        PdfColorFunctionResolutionContext? functionResolutionContext,
        out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (depth > MaxColorSpaceNesting) return false;

        PdfObject? resolved = ResolveColorSpaceDeclaration(value);
        if (resolved is PdfName directName) {
            return TryReadStandardColorSpaceName(directName.Name, out colorSpace);
        }

        if (resolved is not PdfArray { Items.Count: > 0 } array ||
            ResolveColorSpaceDeclaration(array.Items[0]) is not PdfName arrayName) {
            return false;
        }

        switch (arrayName.Name) {
            case "Pattern":
                if (array.Items.Count != 2 ||
                    !TryReadExtendedColorSpaceResource(array.Items[1], depth + 1, evaluationBudget, functionResolutionContext, out PdfPageColorSpace patternBase) ||
                    patternBase.Kind == PdfPageColorSpaceKind.Pattern) return false;
                colorSpace = PdfPageColorSpace.Pattern(patternBase);
                return true;
            case "ICCBased":
                return TryReadIccColorSpace(array, depth, evaluationBudget, functionResolutionContext, out colorSpace);
            case "Indexed":
            case "I":
                return TryReadIndexedColorSpace(array, depth, evaluationBudget, functionResolutionContext, out colorSpace);
            case "Separation":
                return TryReadAlternateColorSpace(array, PdfPageColorSpaceKind.Separation, 1, depth, evaluationBudget, functionResolutionContext, out colorSpace);
            case "DeviceN":
            case "NChannel":
                int componentCount = TryReadDeviceNComponentCount(array);
                return componentCount > 0 &&
                    TryReadAlternateColorSpace(array, PdfPageColorSpaceKind.DeviceN, componentCount, depth, evaluationBudget, functionResolutionContext, out colorSpace);
            case "CalRGB":
                return array.Items.Count > 1 &&
                    ResolveColorSpaceDeclaration(array.Items[1]) is PdfDictionary calibration &&
                    TryReadCalRgbColorSpace(calibration, out colorSpace);
            case "CalGray":
                return array.Items.Count > 1 &&
                    ResolveColorSpaceDeclaration(array.Items[1]) is PdfDictionary grayCalibration &&
                    TryReadCalGrayColorSpace(grayCalibration, out colorSpace);
            case "Lab":
                return array.Items.Count > 1 &&
                    ResolveColorSpaceDeclaration(array.Items[1]) is PdfDictionary labCalibration &&
                    TryReadLabColorSpace(labCalibration, out colorSpace);
            default:
                return TryReadStandardColorSpaceName(arrayName.Name, out colorSpace);
        }
    }

    private bool TryReadIccColorSpace(
        PdfArray array,
        int depth,
        Func<int, bool>? evaluationBudget,
        PdfColorFunctionResolutionContext? functionResolutionContext,
        out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (array.Items.Count < 2) return false;
        PdfObject? resolvedProfile = ResolveColorSpaceDeclaration(array.Items[1]);
        PdfDictionary? profile = resolvedProfile switch {
            PdfStream stream => stream.Dictionary,
            PdfDictionary dictionary => dictionary,
            _ => null
        };
        int? components = profile == null
            ? null
            : TryReadInteger(profile.Items.TryGetValue("N", out PdfObject? count) ? count : null);
        PdfPageColorSpaceKind kind = components switch {
            1 => PdfPageColorSpaceKind.DeviceGray,
            3 => PdfPageColorSpaceKind.DeviceRgb,
            4 => PdfPageColorSpaceKind.DeviceCmyk,
            _ => PdfPageColorSpaceKind.Pattern
        };
        if (kind == PdfPageColorSpaceKind.Pattern) return false;
        int componentCount = components.GetValueOrDefault();
        IReadOnlyList<double>? ranges = null;
        if (profile != null && profile.Items.TryGetValue("Range", out PdfObject? rangeObject)) {
            PdfObject? resolvedRange = ResolveColorSpaceDeclaration(rangeObject);
            if (resolvedRange == null) return false;
            if (resolvedRange is not PdfNull) {
                if (!TryReadIccRange(resolvedRange, componentCount, out ranges)) return false;
                for (int index = 0; index < componentCount; index++) {
                    double minimum = ranges[index * 2];
                    double maximum = ranges[index * 2 + 1];
                    if (!IsFinite(minimum) || !IsFinite(maximum) || minimum > maximum) return false;
                }
            }
        }

        if (resolvedProfile is PdfStream profileStream) {
            if (PdfIccProfileCache.TryRead(profileStream, _objects, _limits.MaxDecodedStreamBytes, out OfficeIccColorProfile? parsedProfile) &&
                parsedProfile != null && parsedProfile.ComponentCount == components) {
                colorSpace = PdfPageColorSpace.IccBased(parsedProfile, ranges);
                return true;
            }
        }

        if (profile != null && profile.Items.TryGetValue("Alternate", out PdfObject? alternateObject)) {
            PdfObject? resolvedAlternate = ResolveColorSpaceDeclaration(alternateObject);
            if (resolvedAlternate == null) return false;
            if (resolvedAlternate is not PdfNull) {
                if (!TryReadExtendedColorSpaceResource(resolvedAlternate, depth + 1, evaluationBudget, functionResolutionContext, out PdfPageColorSpace alternate) ||
                    alternate.Kind == PdfPageColorSpaceKind.Pattern ||
                    alternate.ComponentCount != components) {
                    return false;
                }
                colorSpace = PdfPageColorSpace.IccFallback(alternate, ranges);
                return true;
            }
        }

        colorSpace = PdfPageColorSpace.IccFallback(kind, ranges);
        return true;
    }

    private PdfObject? ResolveColorSpaceDeclaration(PdfObject? value) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        PdfObject? resolved = value;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject indirect)) return null;
            resolved = indirect.Value;
        }
        return resolved;
    }

    private double[] ReadColorSpaceNumberArray(PdfObject? value) {
        if (ResolveColorSpaceDeclaration(value) is not PdfArray { Items.Count: > 0 } array) {
            return Array.Empty<double>();
        }

        var values = new double[array.Items.Count];
        for (int index = 0; index < values.Length; index++) {
            if (ResolveColorSpaceDeclaration(array.Items[index]) is not PdfNumber number) {
                return Array.Empty<double>();
            }
            values[index] = number.Value;
        }
        return values;
    }

    private bool TryReadIccRange(PdfObject value, int componentCount, out IReadOnlyList<double> ranges) {
        ranges = Array.Empty<double>();
        if (value is not PdfArray array || array.Items.Count != componentCount * 2) return false;
        var values = new double[array.Items.Count];
        for (int index = 0; index < values.Length; index++) {
            if (ResolveColorSpaceDeclaration(array.Items[index]) is not PdfNumber number) return false;
            values[index] = number.Value;
        }
        ranges = values;
        return true;
    }

    private bool TryResolveOptionalColorSpaceEntry(
        PdfDictionary dictionary,
        string key,
        out PdfObject? value,
        out bool hasValue) {
        value = null;
        hasValue = false;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? declaration)) return true;
        value = ResolveColorSpaceDeclaration(declaration);
        if (value == null) return false;
        hasValue = value is not PdfNull;
        return true;
    }

    private bool TryReadIndexedColorSpace(
        PdfArray array,
        int depth,
        Func<int, bool>? evaluationBudget,
        PdfColorFunctionResolutionContext? functionResolutionContext,
        out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (array.Items.Count < 4 ||
            !TryReadExtendedColorSpaceResource(array.Items[1], depth + 1, evaluationBudget, functionResolutionContext, out PdfPageColorSpace baseColorSpace) ||
            baseColorSpace.Kind is PdfPageColorSpaceKind.Pattern or PdfPageColorSpaceKind.Indexed ||
            TryReadInteger(array.Items[2]) is not int highValue ||
            highValue < 0 || highValue > 255) {
            return false;
        }

        int componentCount = baseColorSpace.ComponentCount;
        int paletteCount = highValue + 1;
        int lookupLength = checked(paletteCount * componentCount);
        int lookupLimit = Math.Min(lookupLength, _limits.MaxDecodedStreamBytes);
        if (!PdfIndexedImageNormalizer.TryReadIndexedLookupBytes(
                array.Items[3],
                _objects,
                lookupLimit,
                out byte[] lookupBytes)) return false;
        if (lookupBytes.Length < paletteCount * componentCount) return false;

        var lookupComponents = new IReadOnlyList<double>[paletteCount];
        for (int entry = 0; entry < paletteCount; entry++) {
            var components = new double[componentCount];
            for (int component = 0; component < componentCount; component++) {
                components[component] = baseColorSpace.MapLookupByteToComponent(
                    component,
                    lookupBytes[entry * componentCount + component]);
            }
            lookupComponents[entry] = components;
        }

        colorSpace = PdfPageColorSpace.Indexed(baseColorSpace, lookupComponents);
        return true;
    }

    private bool TryReadAlternateColorSpace(
        PdfArray array,
        PdfPageColorSpaceKind kind,
        int componentCount,
        int depth,
        Func<int, bool>? evaluationBudget,
        PdfColorFunctionResolutionContext? functionResolutionContext,
        out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (array.Items.Count < 4 || componentCount < 1 || componentCount > MaxDeviceNComponents ||
            (kind == PdfPageColorSpaceKind.Separation &&
             (ResolveColorSpaceDeclaration(array.Items[1]) is not PdfName colorant ||
              string.Equals(colorant.Name, "None", StringComparison.Ordinal))) ||
            !TryReadExtendedColorSpaceResource(array.Items[2], depth + 1, evaluationBudget, functionResolutionContext, out PdfPageColorSpace alternate) ||
            alternate.Kind is PdfPageColorSpaceKind.Pattern or PdfPageColorSpaceKind.Indexed ||
            !PdfColorSpaceFunctionResolver.TryCreateTintTransform(
                array.Items[3],
                componentCount,
                alternate.ComponentCount,
                _objects,
                _limits.MaxDecodedStreamBytes,
                functionResolutionContext,
                out PdfColorSpaceTintTransform transform,
                out int evaluationCost)) {
            return false;
        }

        colorSpace = PdfPageColorSpace.Alternate(
            kind,
            componentCount,
            alternate,
            transform,
            evaluationCost,
            evaluationBudget);
        return true;
    }

    private int TryReadDeviceNComponentCount(PdfArray array) {
        if (array.Items.Count < 2 || ResolveColorSpaceDeclaration(array.Items[1]) is not PdfArray names) return 0;
        if (names.Items.Count < 1 || names.Items.Count > MaxDeviceNComponents) return 0;
        return names.Items.All(item => ResolveColorSpaceDeclaration(item) is PdfName) ? names.Items.Count : 0;
    }

    private bool TryReadCalGrayColorSpace(PdfDictionary calibration, out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (!PdfCalibratedColorSpaceSemantics.HasSupportedBlackPoint(calibration, _objects) ||
            !calibration.Items.TryGetValue("WhitePoint", out PdfObject? whitePointObject)) return false;
        double[] whitePoint = ReadColorSpaceNumberArray(whitePointObject);
        if (!PdfCalibratedColorSpaceSemantics.IsValidWhitePoint(whitePoint)) return false;

        if (!TryResolveOptionalColorSpaceEntry(calibration, "Gamma", out PdfObject? gammaObject, out bool hasGamma)) return false;
        double gamma = 1D;
        if (hasGamma) {
            if (gammaObject is not PdfNumber gammaNumber ||
                !IsFinite(gammaNumber.Value) || gammaNumber.Value <= 0D) return false;
            gamma = gammaNumber.Value;
        }

        colorSpace = PdfPageColorSpace.CalGray(whitePoint[0], whitePoint[1], whitePoint[2], gamma);
        return true;
    }

    private bool TryReadLabColorSpace(PdfDictionary calibration, out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (!PdfCalibratedColorSpaceSemantics.HasSupportedBlackPoint(calibration, _objects) ||
            !calibration.Items.TryGetValue("WhitePoint", out PdfObject? whitePointObject)) return false;
        double[] whitePoint = ReadColorSpaceNumberArray(whitePointObject);
        if (!PdfCalibratedColorSpaceSemantics.IsValidWhitePoint(whitePoint)) return false;

        if (!TryResolveOptionalColorSpaceEntry(calibration, "Range", out PdfObject? rangeObject, out bool hasRange)) return false;
        IReadOnlyList<double> abRange = new[] { -100D, 100D, -100D, 100D };
        if (hasRange) {
            abRange = ReadColorSpaceNumberArray(rangeObject);
            if (!PdfCalibratedColorSpaceSemantics.IsSupportedLabRange(abRange)) return false;
        }

        colorSpace = PdfPageColorSpace.Lab(whitePoint[0], whitePoint[1], whitePoint[2], abRange);
        return true;
    }

}
