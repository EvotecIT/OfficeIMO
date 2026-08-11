using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Resolves image sample color spaces that can be normalized to managed PNG output.
/// This is also the canonical converter for Indexed palette base color spaces.
/// </summary>
internal sealed class PdfImageColorSpaceNormalization {
    private const int MaxColorSpaceNesting = 8;
    private readonly PdfPageColorSpace _colorSpace;
    private readonly OfficeIccColorProfile? _iccProfile;
    private readonly PdfImageColorSpaceNormalization? _alternateNormalization;
    private readonly Func<IReadOnlyList<double>, IReadOnlyList<double>?>? _tintTransform;
    private readonly double[] _componentRanges;
    private readonly int _cubicEvaluationCost;
    private readonly long _maximumCubicEvaluationWork;

    private PdfImageColorSpaceNormalization(
        PdfPageColorSpace colorSpace,
        int pngColorType,
        OfficeIccColorProfile? iccProfile = null,
        double[]? componentRanges = null,
        PdfImageColorSpaceNormalization? alternateNormalization = null,
        Func<IReadOnlyList<double>, IReadOnlyList<double>?>? tintTransform = null,
        int? sourceColorCount = null,
        bool usesIccApproximation = false,
        int cubicEvaluationCost = 0,
        long maximumCubicEvaluationWork = 0L) {
        _colorSpace = usesIccApproximation ? PdfPageColorSpace.IccFallback(colorSpace) : colorSpace;
        SourceColorCount = sourceColorCount ?? iccProfile?.ComponentCount ?? colorSpace.ComponentCount;
        PngColorType = pngColorType;
        _iccProfile = iccProfile;
        _alternateNormalization = alternateNormalization;
        _tintTransform = tintTransform;
        _componentRanges = componentRanges ?? CreateUnitRanges(SourceColorCount);
        _cubicEvaluationCost = Math.Max(0, cubicEvaluationCost);
        _maximumCubicEvaluationWork = Math.Max(0L, maximumCubicEvaluationWork);
        UsesIccApproximation = usesIccApproximation;
    }

    internal int SourceColorCount { get; }

    internal int PngColorType { get; }

    internal PdfPageColorSpaceKind Kind => _colorSpace.Kind;

    internal bool RequiresColorConversion => _iccProfile != null || _alternateNormalization != null ||
        _colorSpace.Kind is PdfPageColorSpaceKind.CalGray or PdfPageColorSpaceKind.CalRgb or
        PdfPageColorSpaceKind.Lab or PdfPageColorSpaceKind.Indexed;

    internal bool UsesIccApproximation { get; }

    internal bool CanConvertPixelCount(long pixelCount) =>
        pixelCount >= 0L &&
        (_cubicEvaluationCost == 0 ||
         (_maximumCubicEvaluationWork > 0L && pixelCount <= _maximumCubicEvaluationWork / _cubicEvaluationCost));

    internal double[] CreateComponentBuffer() => new double[SourceColorCount];

    internal bool TryConvertPixel(
        byte[] samples,
        int offset,
        PdfImageDecodeTransform? decodeTransform,
        double[] components,
        out OfficeColor color) {
        color = OfficeColor.Black;
        if (components == null || components.Length < SourceColorCount ||
            offset < 0 || offset > samples.Length - SourceColorCount) return false;
        for (int component = 0; component < SourceColorCount; component++) {
            double value = decodeTransform == null
                ? samples[offset + component] / 255D
                : decodeTransform.TransformColorComponentValue(samples[offset + component], component);
            if (decodeTransform == null) value = MapUnitToComponentRange(component, value);
            components[component] = value;
        }
        return TryConvertComponents(components, out color);
    }

    internal bool TryConvertComponents(IReadOnlyList<double> components, out OfficeColor color) {
        color = OfficeColor.Black;
        if (components == null || components.Count < SourceColorCount) return false;
        if (_iccProfile != null) {
            double[] normalized = components as double[] ?? new double[SourceColorCount];
            for (int index = 0; index < SourceColorCount; index++) {
                double minimum = _componentRanges[index * 2];
                double maximum = _componentRanges[index * 2 + 1];
                double value = components[index];
                normalized[index] = value <= minimum ? 0D : value >= maximum ? 1D : (value - minimum) / (maximum - minimum);
            }
            return _iccProfile.TryConvert(normalized, out color);
        }
        if (_alternateNormalization != null) {
            double[] clippedComponents = ClipComponentsToRanges(components);
            IReadOnlyList<double>? alternateComponents = _tintTransform == null
                ? clippedComponents
                : _tintTransform(clippedComponents);
            return alternateComponents != null && _alternateNormalization.TryConvertComponents(alternateComponents, out color);
        }
        if (_colorSpace.Kind == PdfPageColorSpaceKind.DeviceCmyk) {
            byte cyan = ToByte(components[0]);
            byte magenta = ToByte(components[1]);
            byte yellow = ToByte(components[2]);
            byte black = ToByte(components[3]);
            color = OfficeColor.FromRgb(
                ConvertDeviceCmykComponentToRgb(cyan, black),
                ConvertDeviceCmykComponentToRgb(magenta, black),
                ConvertDeviceCmykComponentToRgb(yellow, black));
            return true;
        }
        return _colorSpace.TryConvertColor(components, out color);
    }

    internal double MapLookupByteToComponent(int component, byte value) =>
        MapUnitToComponentRange(component, value / 255D);

    private double[] ClipComponentsToRanges(IReadOnlyList<double> components) {
        double[] clipped = components as double[] ?? new double[SourceColorCount];
        for (int index = 0; index < SourceColorCount; index++) {
            double minimum = _componentRanges[index * 2];
            double maximum = _componentRanges[index * 2 + 1];
            double value = components[index];
            clipped[index] = value < minimum ? minimum : value > maximum ? maximum : value;
        }
        return clipped;
    }

    internal static bool TryResolve(
        PdfObject? colorSpaceObj,
        string colorSpaceName,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        out PdfImageColorSpaceNormalization normalization) =>
        TryResolve(colorSpaceObj, colorSpaceName, objects, maxDecodedStreamBytes, depth: 0, out normalization);

    private static bool TryResolve(
        PdfObject? colorSpaceObj,
        string colorSpaceName,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int depth,
        out PdfImageColorSpaceNormalization normalization) {
        normalization = null!;
        if (depth > MaxColorSpaceNesting || maxDecodedStreamBytes <= 0) return false;

        PdfObject? resolvedColorSpace = ResolveObject(colorSpaceObj, objects);
        if (resolvedColorSpace is PdfName directName) colorSpaceName = directName.Name;
        if (TryCreateFromNamedColorSpace(colorSpaceName, out normalization)) return true;

        if (resolvedColorSpace is not PdfArray { Items.Count: > 0 } colorSpaceArray ||
            ResolveObject(colorSpaceArray.Items[0], objects) is not PdfName colorSpaceKind) return false;

        if (TryCreateFromNamedColorSpace(colorSpaceKind.Name, out normalization)) return true;
        switch (colorSpaceKind.Name) {
            case "CalRGB":
                return TryCreateCalRgb(colorSpaceArray, objects, out normalization);
            case "CalGray":
                return TryCreateCalGray(colorSpaceArray, objects, out normalization);
            case "Lab":
                return TryCreateLab(colorSpaceArray, objects, out normalization);
            case "ICCBased":
            case "ICC":
                return TryCreateIcc(
                    colorSpaceArray,
                    objects,
                    maxDecodedStreamBytes,
                    depth,
                    out normalization);
            case "Indexed":
            case "I":
                return TryCreateIndexed(colorSpaceArray, objects, maxDecodedStreamBytes, depth, out normalization);
            case "Separation":
                return TryCreateSpecial(colorSpaceArray, PdfPageColorSpaceKind.Separation, 1, objects, maxDecodedStreamBytes, depth, out normalization);
            case "DeviceN":
            case "NChannel":
                if (colorSpaceArray.Items.Count < 2 ||
                    ResolveObject(colorSpaceArray.Items[1], objects) is not PdfArray names ||
                    names.Items.Count < 1 || names.Items.Count > 32 ||
                    names.Items.Any(item => ResolveObject(item, objects) is not PdfName)) return false;
                return TryCreateSpecial(colorSpaceArray, PdfPageColorSpaceKind.DeviceN, names.Items.Count, objects, maxDecodedStreamBytes, depth, out normalization);
            default:
                return false;
        }
    }

    private static bool TryCreateIcc(
        PdfArray colorSpaceArray,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int depth,
        out PdfImageColorSpaceNormalization normalization) {
        normalization = null!;
        if (colorSpaceArray.Items.Count < 2 ||
            ResolveObject(colorSpaceArray.Items[1], objects) is not PdfStream profileStream ||
            TryReadIccComponentCount(profileStream, objects) is not int componentCount) return false;

        double[]? ranges = TryReadIccRanges(profileStream.Dictionary, componentCount, objects, out bool rangeIsValid);
        if (!rangeIsValid) return false;
        if (PdfIccProfileCache.TryRead(profileStream, objects, maxDecodedStreamBytes, out OfficeIccColorProfile? profile) &&
            profile != null && profile.ComponentCount == componentCount) {
            PdfPageColorSpace output = profile.ComponentCount == 1
                ? PdfPageColorSpaceKind.DeviceGray
                : PdfPageColorSpaceKind.DeviceRgb;
            normalization = new PdfImageColorSpaceNormalization(output, 2, profile, ranges);
            return true;
        }

        if (profileStream.Dictionary.Items.TryGetValue("Alternate", out PdfObject? alternateObject)) {
            string alternateName = ResolveObject(alternateObject, objects) is PdfName name ? name.Name : string.Empty;
            if (!TryResolve(alternateObject, alternateName, objects, maxDecodedStreamBytes, depth + 1, out PdfImageColorSpaceNormalization alternate) ||
                alternate.SourceColorCount != componentCount) return false;
            normalization = new PdfImageColorSpaceNormalization(
                alternate._colorSpace,
                alternate.PngColorType,
                componentRanges: ranges,
                alternateNormalization: alternate,
                sourceColorCount: componentCount,
                usesIccApproximation: true);
            return true;
        }

        PdfPageColorSpace fallback = componentCount switch {
            1 => PdfPageColorSpaceKind.DeviceGray,
            3 => PdfPageColorSpaceKind.DeviceRgb,
            4 => PdfPageColorSpaceKind.DeviceCmyk,
            _ => PdfPageColorSpaceKind.Pattern
        };
        if (fallback.Kind == PdfPageColorSpaceKind.Pattern) return false;
        normalization = new PdfImageColorSpaceNormalization(
            fallback,
            componentCount == 1 ? 0 : 2,
            componentRanges: ranges,
            usesIccApproximation: true);
        return true;
    }

    private static bool TryCreateFromNamedColorSpace(string colorSpaceName, out PdfImageColorSpaceNormalization normalization) {
        switch (colorSpaceName) {
            case "DeviceGray":
            case "G":
                normalization = new PdfImageColorSpaceNormalization(PdfPageColorSpaceKind.DeviceGray, 0);
                return true;
            case "DeviceRGB":
            case "RGB":
                normalization = new PdfImageColorSpaceNormalization(PdfPageColorSpaceKind.DeviceRgb, 2);
                return true;
            case "DeviceCMYK":
            case "CMYK":
                normalization = new PdfImageColorSpaceNormalization(PdfPageColorSpaceKind.DeviceCmyk, 2);
                return true;
            default:
                normalization = null!;
                return false;
        }
    }

    private static bool TryCreateCalRgb(
        PdfArray colorSpaceArray,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageColorSpaceNormalization normalization) {
        normalization = null!;
        if (colorSpaceArray.Items.Count < 2 ||
            ResolveObject(colorSpaceArray.Items[1], objects) is not PdfDictionary calibration ||
            !TryReadNumberArray(calibration, "WhitePoint", 3, objects, out double[] whitePoint) ||
            whitePoint.Any(value => value <= 0D)) return false;

        double[]? gamma = null;
        if (calibration.Items.ContainsKey("Gamma") &&
            (!TryReadNumberArray(calibration, "Gamma", 3, objects, out gamma) || gamma.Any(value => value <= 0D))) return false;
        double[]? matrix = null;
        if (calibration.Items.ContainsKey("Matrix") &&
            !TryReadNumberArray(calibration, "Matrix", 9, objects, out matrix)) return false;

        normalization = new PdfImageColorSpaceNormalization(
            PdfPageColorSpace.CalRgb(whitePoint[0], whitePoint[1], whitePoint[2], gamma, matrix),
            2);
        return true;
    }

    private static bool TryCreateIndexed(
        PdfArray colorSpaceArray,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int depth,
        out PdfImageColorSpaceNormalization normalization) {
        normalization = null!;
        if (colorSpaceArray.Items.Count < 4 ||
            !TryReadInteger(colorSpaceArray.Items[2], objects, out int highValue) ||
            highValue < 0 || highValue > 255) return false;
        PdfObject? baseObject = colorSpaceArray.Items[1];
        string baseName = ResolveObject(baseObject, objects) is PdfName name ? name.Name : string.Empty;
        if (!TryResolve(baseObject, baseName, objects, maxDecodedStreamBytes, depth + 1, out PdfImageColorSpaceNormalization baseColorSpace) ||
            baseColorSpace.Kind is PdfPageColorSpaceKind.Indexed or PdfPageColorSpaceKind.Pattern) return false;
        int paletteCount = highValue + 1;
        int lookupLength = checked(paletteCount * baseColorSpace.SourceColorCount);
        if (!PdfIndexedImageNormalizer.TryReadIndexedLookupBytes(
                colorSpaceArray.Items[3],
                objects,
                Math.Min(lookupLength, maxDecodedStreamBytes),
                out byte[] lookup) || lookup.Length < lookupLength) return false;
        var palette = new OfficeColor[paletteCount];
        var components = new double[baseColorSpace.SourceColorCount];
        for (int entry = 0; entry < paletteCount; entry++) {
            for (int component = 0; component < components.Length; component++) {
                components[component] = baseColorSpace.MapLookupByteToComponent(
                    component,
                    lookup[entry * components.Length + component]);
            }
            if (!baseColorSpace.TryConvertComponents(components, out palette[entry])) return false;
        }
        normalization = new PdfImageColorSpaceNormalization(
            PdfPageColorSpace.Indexed(palette, baseColorSpace.UsesIccApproximation),
            2,
            componentRanges: new[] { 0D, (double)highValue });
        return true;
    }

    private static bool TryCreateSpecial(
        PdfArray colorSpaceArray,
        PdfPageColorSpaceKind kind,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int depth,
        out PdfImageColorSpaceNormalization normalization) {
        normalization = null!;
        if (colorSpaceArray.Items.Count < 4) return false;
        PdfObject? alternateObject = colorSpaceArray.Items[2];
        string alternateName = ResolveObject(alternateObject, objects) is PdfName name ? name.Name : string.Empty;
        if (!TryResolve(alternateObject, alternateName, objects, maxDecodedStreamBytes, depth + 1, out PdfImageColorSpaceNormalization alternate) ||
            !PdfColorSpaceFunctionResolver.TryCreateTintTransform(
                colorSpaceArray.Items[3],
                componentCount,
                alternate.SourceColorCount,
                objects,
                maxDecodedStreamBytes,
                out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform,
                out int cubicEvaluationCost)) return false;
        int aggregateCubicCost = checked(cubicEvaluationCost + alternate._cubicEvaluationCost);
        normalization = new PdfImageColorSpaceNormalization(
            kind,
            2,
            alternateNormalization: alternate,
            tintTransform: transform,
            sourceColorCount: componentCount,
            usesIccApproximation: alternate.UsesIccApproximation,
            cubicEvaluationCost: aggregateCubicCost,
            maximumCubicEvaluationWork: Math.Max(1L, maxDecodedStreamBytes / sizeof(double)));
        return true;
    }

    private static bool TryCreateCalGray(
        PdfArray colorSpaceArray,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageColorSpaceNormalization normalization) {
        normalization = null!;
        if (colorSpaceArray.Items.Count < 2 ||
            ResolveObject(colorSpaceArray.Items[1], objects) is not PdfDictionary calibration ||
            !TryReadNumberArray(calibration, "WhitePoint", 3, objects, out double[] whitePoint) ||
            whitePoint.Any(value => value <= 0D)) return false;
        double gamma = 1D;
        if (calibration.Items.TryGetValue("Gamma", out PdfObject? gammaObject)) {
            if (ResolveObject(gammaObject, objects) is not PdfNumber gammaNumber ||
                double.IsNaN(gammaNumber.Value) || double.IsInfinity(gammaNumber.Value) || gammaNumber.Value <= 0D) return false;
            gamma = gammaNumber.Value;
        }
        normalization = new PdfImageColorSpaceNormalization(
            PdfPageColorSpace.CalGray(whitePoint[0], whitePoint[1], whitePoint[2], gamma),
            0);
        return true;
    }

    private static bool TryCreateLab(
        PdfArray colorSpaceArray,
        Dictionary<int, PdfIndirectObject> objects,
        out PdfImageColorSpaceNormalization normalization) {
        normalization = null!;
        if (colorSpaceArray.Items.Count < 2 ||
            ResolveObject(colorSpaceArray.Items[1], objects) is not PdfDictionary calibration ||
            !TryReadNumberArray(calibration, "WhitePoint", 3, objects, out double[] whitePoint) ||
            whitePoint.Any(value => value <= 0D)) return false;
        double[] abRange = { -100D, 100D, -100D, 100D };
        if (calibration.Items.ContainsKey("Range") &&
            (!TryReadNumberArray(calibration, "Range", 4, objects, out abRange) ||
             abRange[0] >= abRange[1] || abRange[2] >= abRange[3])) return false;
        double[] componentRanges = { 0D, 100D, abRange[0], abRange[1], abRange[2], abRange[3] };
        normalization = new PdfImageColorSpaceNormalization(
            PdfPageColorSpace.Lab(whitePoint[0], whitePoint[1], whitePoint[2], abRange),
            2,
            componentRanges: componentRanges);
        return true;
    }

    private static bool TryReadNumberArray(
        PdfDictionary dictionary,
        string key,
        int count,
        Dictionary<int, PdfIndirectObject> objects,
        out double[] values) {
        values = Array.Empty<double>();
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(value, objects) is not PdfArray array || array.Items.Count != count) return false;
        values = new double[count];
        for (int index = 0; index < count; index++) {
            if (ResolveObject(array.Items[index], objects) is not PdfNumber number ||
                double.IsNaN(number.Value) || double.IsInfinity(number.Value)) return false;
            values[index] = number.Value;
        }
        return true;
    }

    private static bool TryReadInteger(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        out int number) {
        number = 0;
        if (ResolveObject(value, objects) is not PdfNumber pdfNumber ||
            double.IsNaN(pdfNumber.Value) || double.IsInfinity(pdfNumber.Value) ||
            Math.Truncate(pdfNumber.Value) != pdfNumber.Value ||
            pdfNumber.Value < int.MinValue || pdfNumber.Value > int.MaxValue) return false;
        number = (int)pdfNumber.Value;
        return true;
    }

    private static int? TryReadIccComponentCount(PdfStream profileStream, Dictionary<int, PdfIndirectObject> objects) {
        if (!profileStream.Dictionary.Items.TryGetValue("N", out PdfObject? countObject) ||
            ResolveObject(countObject, objects) is not PdfNumber countNumber ||
            countNumber.Value < 0 || countNumber.Value > int.MaxValue ||
            Math.Truncate(countNumber.Value) != countNumber.Value) return null;
        return (int)countNumber.Value;
    }

    private static double[]? TryReadIccRanges(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out bool isValid) {
        isValid = true;
        if (!dictionary.Items.TryGetValue("Range", out PdfObject? rangeObject)) return null;
        if (ResolveObject(rangeObject, objects) is not PdfArray range || range.Items.Count != componentCount * 2) {
            isValid = false;
            return null;
        }
        var values = new double[range.Items.Count];
        for (int index = 0; index < values.Length; index += 2) {
            if (ResolveObject(range.Items[index], objects) is not PdfNumber minimum ||
                ResolveObject(range.Items[index + 1], objects) is not PdfNumber maximum ||
                double.IsNaN(minimum.Value) || double.IsInfinity(minimum.Value) ||
                double.IsNaN(maximum.Value) || double.IsInfinity(maximum.Value) ||
                minimum.Value >= maximum.Value) {
                isValid = false;
                return null;
            }
            values[index] = minimum.Value;
            values[index + 1] = maximum.Value;
        }
        return values;
    }

    private static PdfObject? ResolveObject(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) =>
        PdfObjectLookup.Resolve(objects, obj);

    private double MapUnitToComponentRange(int component, double value) {
        double minimum = _componentRanges[component * 2];
        double maximum = _componentRanges[component * 2 + 1];
        return minimum + value * (maximum - minimum);
    }

    private static double[] CreateUnitRanges(int componentCount) {
        var ranges = new double[componentCount * 2];
        for (int component = 0; component < componentCount; component++) ranges[component * 2 + 1] = 1D;
        return ranges;
    }

    private static byte ToByte(double value) =>
        (byte)Math.Round((value < 0D ? 0D : value > 1D ? 1D : value) * 255D);

    private static byte ConvertDeviceCmykComponentToRgb(byte colorant, byte black) {
        int ink = colorant + black;
        return (byte)(255 - (ink > 255 ? 255 : ink));
    }
}
