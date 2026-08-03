using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private const int MaxColorSpaceNesting = 8;
    private const int MaxDeviceNComponents = 32;

    private bool TryReadExtendedColorSpaceResource(PdfObject? value, int depth, out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (depth > MaxColorSpaceNesting) return false;

        PdfObject? resolved = ResolveObject(value);
        if (resolved is PdfName directName) {
            return TryReadStandardColorSpaceName(directName.Name, out colorSpace);
        }

        if (resolved is not PdfArray { Items.Count: > 0 } array ||
            ResolveObject(array.Items[0]) is not PdfName arrayName) {
            return false;
        }

        switch (arrayName.Name) {
            case "ICCBased":
                return TryReadIccColorSpace(array, out colorSpace);
            case "Indexed":
            case "I":
                return TryReadIndexedColorSpace(array, depth, out colorSpace);
            case "Separation":
                return TryReadAlternateColorSpace(array, PdfPageColorSpaceKind.Separation, 1, depth, out colorSpace);
            case "DeviceN":
            case "NChannel":
                int componentCount = TryReadDeviceNComponentCount(array);
                return componentCount > 0 &&
                    TryReadAlternateColorSpace(array, PdfPageColorSpaceKind.DeviceN, componentCount, depth, out colorSpace);
            case "CalRGB":
                return array.Items.Count > 1 &&
                    ResolveDictionary(array.Items[1]) is PdfDictionary calibration &&
                    TryReadCalRgbColorSpace(calibration, out colorSpace);
            default:
                return TryReadStandardColorSpaceName(arrayName.Name, out colorSpace);
        }
    }

    private bool TryReadIccColorSpace(PdfArray array, out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (array.Items.Count < 2) return false;
        PdfDictionary? profile = ResolveObject(array.Items[1]) switch {
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
        colorSpace = PdfPageColorSpace.IccBased(kind);
        return true;
    }

    private bool TryReadIndexedColorSpace(PdfArray array, int depth, out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (array.Items.Count < 4 ||
            !TryReadExtendedColorSpaceResource(array.Items[1], depth + 1, out PdfPageColorSpace baseColorSpace) ||
            baseColorSpace.Kind is PdfPageColorSpaceKind.Pattern or PdfPageColorSpaceKind.Indexed ||
            TryReadInteger(array.Items[2]) is not int highValue ||
            highValue < 0 || highValue > 255 ||
            !PdfIndexedImageNormalizer.TryReadIndexedLookupBytes(array.Items[3], _objects, out byte[] lookupBytes)) {
            return false;
        }

        int componentCount = baseColorSpace.ComponentCount;
        int paletteCount = highValue + 1;
        if (lookupBytes.Length < paletteCount * componentCount) return false;

        var palette = new OfficeColor[paletteCount];
        var components = new double[componentCount];
        for (int entry = 0; entry < paletteCount; entry++) {
            for (int component = 0; component < componentCount; component++) {
                components[component] = lookupBytes[entry * componentCount + component] / 255D;
            }
            if (!baseColorSpace.TryConvertColor(components, out palette[entry])) return false;
        }

        colorSpace = PdfPageColorSpace.Indexed(palette, baseColorSpace.UsesIccApproximation);
        return true;
    }

    private bool TryReadAlternateColorSpace(
        PdfArray array,
        PdfPageColorSpaceKind kind,
        int componentCount,
        int depth,
        out PdfPageColorSpace colorSpace) {
        colorSpace = PdfPageColorSpaceKind.DeviceGray;
        if (array.Items.Count < 4 || componentCount < 1 || componentCount > MaxDeviceNComponents ||
            !TryReadExtendedColorSpaceResource(array.Items[2], depth + 1, out PdfPageColorSpace alternate) ||
            alternate.Kind is PdfPageColorSpaceKind.Pattern or PdfPageColorSpaceKind.Indexed ||
            !TryCreateTintTransform(array.Items[3], componentCount, alternate.ComponentCount, out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform)) {
            return false;
        }

        colorSpace = PdfPageColorSpace.Alternate(kind, componentCount, alternate, transform);
        return true;
    }

    private int TryReadDeviceNComponentCount(PdfArray array) {
        if (array.Items.Count < 2 || ResolveObject(array.Items[1]) is not PdfArray names) return 0;
        if (names.Items.Count < 1 || names.Items.Count > MaxDeviceNComponents) return 0;
        return names.Items.All(item => ResolveObject(item) is PdfName) ? names.Items.Count : 0;
    }

    private bool TryCreateTintTransform(
        PdfObject? value,
        int inputCount,
        int outputCount,
        out Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform) {
        transform = null!;
        PdfObject? resolved = ResolveObject(value);
        PdfDictionary? dictionary = resolved switch {
            PdfStream stream => stream.Dictionary,
            PdfDictionary direct => direct,
            _ => null
        };
        if (dictionary == null) return false;

        int? functionType = TryReadInteger(dictionary.Items.TryGetValue("FunctionType", out PdfObject? type) ? type : null);
        if (functionType == 2 && inputCount == 1) {
            IReadOnlyList<double> c0 = dictionary.Items.TryGetValue("C0", out PdfObject? c0Value)
                ? ReadNumberArray(c0Value)
                : new[] { 0D };
            IReadOnlyList<double> c1 = dictionary.Items.TryGetValue("C1", out PdfObject? c1Value)
                ? ReadNumberArray(c1Value)
                : new[] { 1D };
            if (!dictionary.Items.TryGetValue("N", out PdfObject? exponentValue) ||
                ResolveObject(exponentValue) is not PdfNumber exponentNumber) return false;
            double exponent = exponentNumber.Value;
            if (c0.Count != outputCount || c1.Count != outputCount ||
                c0.Any(value => !IsFinite(value)) || c1.Any(value => !IsFinite(value)) ||
                !IsFinite(exponent) || exponent <= 0D ||
                !HasUnitFunctionBounds(dictionary, inputCount, outputCount, requireRange: false)) return false;
            transform = components => EvaluateType2TintTransform(components, c0, c1, exponent);
            return true;
        }

        if (functionType == 4 && resolved is PdfStream calculator &&
            HasUnitFunctionBounds(dictionary, inputCount, outputCount, requireRange: true) &&
            TryReadBoundedCalculatorProgram(calculator, inputCount, outputCount, out int duplicateCount)) {
            transform = components => EvaluateIdentityTintTransform(components, inputCount, outputCount, duplicateCount);
            return true;
        }

        return false;
    }

    private bool HasUnitFunctionBounds(PdfDictionary dictionary, int inputCount, int outputCount, bool requireRange) {
        if (!dictionary.Items.TryGetValue("Domain", out PdfObject? domainObject) ||
            !HasUnitIntervals(ReadNumberArray(domainObject), inputCount)) return false;
        if (!dictionary.Items.TryGetValue("Range", out PdfObject? rangeObject)) return !requireRange;
        return HasUnitIntervals(ReadNumberArray(rangeObject), outputCount);
    }

    private static bool HasUnitIntervals(IReadOnlyList<double> values, int count) {
        if (values.Count != count * 2) return false;
        for (int index = 0; index < count; index++) {
            if (values[index * 2] != 0D || values[index * 2 + 1] != 1D) return false;
        }
        return true;
    }

    private bool TryReadBoundedCalculatorProgram(PdfStream stream, int inputCount, int outputCount, out int duplicateCount) {
        duplicateCount = 0;
        if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects).Count != 0) return false;
        byte[] bytes = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, _objects);
        if (bytes.Length > 256) return false;
        string program = Encoding.ASCII.GetString(bytes).Trim();
        if (program.Length < 2 || program[0] != '{' || program[program.Length - 1] != '}') return false;
        string body = program.Substring(1, program.Length - 2).Trim();
        if (body.Length == 0) return inputCount == outputCount;
        string[] tokens = body.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (inputCount != 1 || tokens.Length > MaxDeviceNComponents || tokens.Any(token => !string.Equals(token, "dup", StringComparison.Ordinal))) return false;
        duplicateCount = tokens.Length;
        return outputCount == duplicateCount + 1;
    }

    private static double[]? EvaluateType2TintTransform(
        IReadOnlyList<double> components,
        IReadOnlyList<double> c0,
        IReadOnlyList<double> c1,
        double exponent) {
        if (components.Count < 1 || !IsFinite(components[0])) return null;
        double factor = Math.Pow(ClampColorComponent(components[0]), exponent);
        var result = new double[c0.Count];
        for (int i = 0; i < result.Length; i++) result[i] = c0[i] + factor * (c1[i] - c0[i]);
        return result;
    }

    private static double[]? EvaluateIdentityTintTransform(
        IReadOnlyList<double> components,
        int inputCount,
        int outputCount,
        int duplicateCount) {
        if (components.Count < inputCount || components.Take(inputCount).Any(value => !IsFinite(value))) return null;
        if (duplicateCount == 0) return components.Take(outputCount).Select(ClampColorComponent).ToArray();
        return Enumerable.Repeat(ClampColorComponent(components[0]), outputCount).ToArray();
    }

    private static double ClampColorComponent(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
}
