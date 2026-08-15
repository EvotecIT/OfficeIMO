using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageColorSpace {
    private static readonly double[] IdentityGamma = { 1D, 1D, 1D };
    private static readonly double[] IdentityMatrix = { 1D, 0D, 0D, 0D, 1D, 0D, 0D, 0D, 1D };
    private readonly PdfPageCalRgbParameters? _calRgb;
    private readonly PdfPageCustomColorSpace? _custom;

    public PdfPageColorSpace(PdfPageColorSpaceKind kind) {
        Kind = kind;
        _calRgb = null;
        _custom = null;
    }

    private PdfPageColorSpace(PdfPageCalRgbParameters calRgb) {
        Kind = PdfPageColorSpaceKind.CalRgb;
        _calRgb = calRgb;
        _custom = null;
    }

    private PdfPageColorSpace(PdfPageColorSpaceKind kind, PdfPageCustomColorSpace custom) {
        Kind = kind;
        _calRgb = null;
        _custom = custom;
    }

    public PdfPageColorSpaceKind Kind { get; }

    public int ComponentCount => _custom?.ComponentCount ?? Kind switch {
        PdfPageColorSpaceKind.DeviceRgb or PdfPageColorSpaceKind.CalRgb or PdfPageColorSpaceKind.Lab => 3,
        PdfPageColorSpaceKind.DeviceCmyk => 4,
        _ => 1
    };

    public bool UsesIccApproximation => _custom?.UsesIccApproximation == true;

    internal bool IsNativeDeviceGray =>
        Kind == PdfPageColorSpaceKind.DeviceGray && _custom == null;

    internal bool IsNativeDeviceRgb =>
        Kind == PdfPageColorSpaceKind.DeviceRgb && _custom == null;

    internal bool IsNativeDeviceCmyk =>
        Kind == PdfPageColorSpaceKind.DeviceCmyk && _custom == null;

    internal bool TryGetOutputProfileComponents(
        IReadOnlyList<double> components,
        int profileComponentCount,
        out IReadOnlyList<double> profileComponents) =>
        TryGetOutputProfileComponents(components, profileComponentCount, depth: 0, out profileComponents);

    private bool TryGetOutputProfileComponents(
        IReadOnlyList<double> components,
        int profileComponentCount,
        int depth,
        out IReadOnlyList<double> profileComponents) {
        profileComponents = Array.Empty<double>();
        if (components == null || depth > 8) return false;
        if ((IsNativeDeviceRgb && profileComponentCount == 3) ||
            (IsNativeDeviceCmyk && profileComponentCount == 4)) {
            if (components.Count < profileComponentCount) return false;
            profileComponents = components;
            return true;
        }
        if (TryGetIndexedBaseComponents(components, out PdfPageColorSpace indexedBase, out IReadOnlyList<double> indexedComponents)) {
            return indexedBase.TryGetOutputProfileComponents(
                indexedComponents,
                profileComponentCount,
                depth + 1,
                out profileComponents);
        }
        if (Kind is not (PdfPageColorSpaceKind.Separation or PdfPageColorSpaceKind.DeviceN) ||
            components.Count < ComponentCount ||
            _custom?.Alternate is not PdfPageColorSpace alternate ||
            _custom.Transform == null) return false;

        var alternateComponents = new double[alternate.ComponentCount];
        return _custom.Transform(components, alternateComponents) &&
            alternate.TryGetOutputProfileComponents(
                alternateComponents,
                profileComponentCount,
                depth + 1,
                out profileComponents);
    }

    internal bool TryGetIndexedBaseComponents(
        IReadOnlyList<double> components,
        out PdfPageColorSpace baseColorSpace,
        out IReadOnlyList<double> baseComponents) {
        baseColorSpace = default;
        baseComponents = Array.Empty<double>();
        if (Kind != PdfPageColorSpaceKind.Indexed ||
            components == null || components.Count == 0 ||
            _custom?.IndexedBaseColorSpace is not PdfPageColorSpace indexedBase ||
            _custom.IndexedLookupComponents is not IReadOnlyList<IReadOnlyList<double>> lookup ||
            lookup.Count == 0) return false;

        int index = (int)Math.Round(components[0]);
        if (index < 0) index = 0;
        if (index >= lookup.Count) index = lookup.Count - 1;
        baseColorSpace = indexedBase;
        baseComponents = lookup[index];
        return true;
    }

    public bool HasPatternBaseColorSpace => Kind == PdfPageColorSpaceKind.Pattern && _custom?.Alternate != null;

    internal bool RequiresColorManagedGradientSampling =>
        _custom is not null ||
        _calRgb is not null ||
        Kind == PdfPageColorSpaceKind.DeviceCmyk;

    public static PdfPageColorSpace CalRgb(
        double whiteX,
        double whiteY,
        double whiteZ,
        IReadOnlyList<double>? gamma,
        IReadOnlyList<double>? matrix) =>
        new PdfPageColorSpace(new PdfPageCalRgbParameters(whiteX, whiteY, whiteZ, gamma, matrix));

    public static PdfPageColorSpace CalGray(double whiteX, double whiteY, double whiteZ, double gamma) =>
        new PdfPageColorSpace(
            PdfPageColorSpaceKind.CalGray,
            new PdfPageCustomColorSpace(
                1,
                (components, _) => OfficeColorSpaceConverter.FromCalibratedGray(
                    components[0], whiteX, whiteY, whiteZ, gamma)));

    public static PdfPageColorSpace Lab(
        double whiteX,
        double whiteY,
        double whiteZ,
        IReadOnlyList<double> abRange) =>
        new PdfPageColorSpace(
            PdfPageColorSpaceKind.Lab,
            new PdfPageCustomColorSpace(
                3,
                (components, _) => OfficeColorSpaceConverter.FromLab(
                    Clamp(components[0], 0D, 100D),
                    Clamp(components[1], abRange[0], abRange[1]),
                    Clamp(components[2], abRange[2], abRange[3]),
                    whiteX,
                    whiteY,
                    whiteZ),
                componentRanges: new[] { 0D, 100D, abRange[0], abRange[1], abRange[2], abRange[3] }));

    public static PdfPageColorSpace IccBased(PdfPageColorSpaceKind alternateKind) =>
        new PdfPageColorSpace(alternateKind, new PdfPageCustomColorSpace(ComponentCountFor(alternateKind), true));

    public static PdfPageColorSpace IccBased(OfficeIccColorProfile profile, IReadOnlyList<double>? ranges = null) =>
        new PdfPageColorSpace(
            profile.ComponentCount == 1 ? PdfPageColorSpaceKind.DeviceGray : PdfPageColorSpaceKind.DeviceRgb,
            new PdfPageCustomColorSpace(
                profile.ComponentCount,
                (components, renderingIntent) => profile.TryConvert(
                    NormalizeIccComponents(components, profile.ComponentCount, ranges),
                    renderingIntent,
                    out OfficeColor color)
                    ? color
                    : (OfficeColor?)null,
                componentRanges: ranges));

    public static PdfPageColorSpace IccFallback(
        PdfPageColorSpace alternate,
        IReadOnlyList<double>? ranges = null) =>
        new PdfPageColorSpace(
            alternate.Kind,
            new PdfPageCustomColorSpace(
                alternate.ComponentCount,
                (components, renderingIntent) => alternate.TryConvertColor(
                    ClipIccComponents(components, alternate.ComponentCount, ranges),
                    renderingIntent,
                    out OfficeColor color)
                    ? color
                    : (OfficeColor?)null,
                usesIccApproximation: true,
                componentRanges: ranges));

    public static PdfPageColorSpace Indexed(IReadOnlyList<OfficeColor> palette, bool usesIccApproximation) =>
        new PdfPageColorSpace(PdfPageColorSpaceKind.Indexed, new PdfPageCustomColorSpace(palette, usesIccApproximation));

    public static PdfPageColorSpace Indexed(
        PdfPageColorSpace baseColorSpace,
        IReadOnlyList<IReadOnlyList<double>> lookupComponents) =>
        new PdfPageColorSpace(
            PdfPageColorSpaceKind.Indexed,
            new PdfPageCustomColorSpace(baseColorSpace, lookupComponents));

    public static PdfPageColorSpace Alternate(
        PdfPageColorSpaceKind kind,
        int componentCount,
        PdfPageColorSpace alternate,
        PdfColorSpaceTintTransform transform,
        int evaluationCost,
        Func<int, bool>? evaluationBudget) =>
        new PdfPageColorSpace(
            kind,
            new PdfPageCustomColorSpace(componentCount, alternate, transform, evaluationCost, evaluationBudget));

    public static PdfPageColorSpace Pattern(PdfPageColorSpace baseColorSpace) =>
        new PdfPageColorSpace(PdfPageColorSpaceKind.Pattern, new PdfPageCustomColorSpace(baseColorSpace));

    public OfficeColor ConvertCalRgb(double red, double green, double blue) {
        PdfPageCalRgbParameters parameters = _calRgb ?? PdfPageCalRgbParameters.Default;
        return OfficeColorSpaceConverter.FromCalibratedRgb(
            red, green, blue,
            parameters.WhiteX, parameters.WhiteY, parameters.WhiteZ,
            parameters.Gamma, parameters.Matrix);
    }

    public bool TryConvertColor(IReadOnlyList<double> components, out OfficeColor color) {
        return TryConvertColor(components, OfficeIccRenderingIntent.RelativeColorimetric, out color);
    }

    public bool TryConvertColor(
        IReadOnlyList<double> components,
        OfficeIccRenderingIntent renderingIntent,
        out OfficeColor color) {
        color = OfficeColor.Black;
        if (components == null || components.Count < ComponentCount || Kind == PdfPageColorSpaceKind.Pattern) return false;
        for (int index = 0; index < ComponentCount; index++) if (!IsFinite(components[index])) return false;

        if (_custom?.ColorTransform != null) {
            OfficeColor? transformed = _custom.ColorTransform(components, renderingIntent);
            if (!transformed.HasValue) return false;
            color = transformed.Value;
            return true;
        }

        if (Kind == PdfPageColorSpaceKind.Indexed) {
            IReadOnlyList<IReadOnlyList<double>>? lookupComponents = _custom?.IndexedLookupComponents;
            if (lookupComponents != null && _custom?.IndexedBaseColorSpace is PdfPageColorSpace indexedBase) {
                if (lookupComponents.Count == 0) return false;
                int dynamicIndex = (int)Math.Round(components[0]);
                if (dynamicIndex < 0) dynamicIndex = 0;
                if (dynamicIndex >= lookupComponents.Count) dynamicIndex = lookupComponents.Count - 1;
                return indexedBase.TryConvertColor(lookupComponents[dynamicIndex], renderingIntent, out color);
            }

            IReadOnlyList<OfficeColor>? palette = _custom?.Palette;
            if (palette == null || palette.Count == 0) return false;
            int index = (int)Math.Round(components[0]);
            if (index < 0) index = 0;
            if (index >= palette.Count) index = palette.Count - 1;
            color = palette[index];
            return true;
        }

        if (Kind is PdfPageColorSpaceKind.Separation or PdfPageColorSpaceKind.DeviceN) {
            if (_custom?.Alternate is not PdfPageColorSpace alternate || _custom.Transform == null) return false;
            var transformed = new double[alternate.ComponentCount];
            return _custom.Transform(components, transformed) && alternate.TryConvertColor(transformed, renderingIntent, out color);
        }

        switch (Kind) {
            case PdfPageColorSpaceKind.DeviceRgb:
                color = OfficeColor.FromRgb(ToByte(components[0]), ToByte(components[1]), ToByte(components[2]));
                return true;
            case PdfPageColorSpaceKind.DeviceCmyk:
                color = OfficeColorSpaceConverter.FromCmyk(components[0], components[1], components[2], components[3]);
                return true;
            case PdfPageColorSpaceKind.CalGray:
                color = PdfPageColorConverter.FromCalGray(components[0]);
                return true;
            case PdfPageColorSpaceKind.CalRgb:
                color = ConvertCalRgb(components[0], components[1], components[2]);
                return true;
            case PdfPageColorSpaceKind.Lab:
                color = PdfPageColorConverter.FromLab(components[0], components[1], components[2]);
                return true;
            default:
                byte gray = ToByte(components[0]);
                color = OfficeColor.FromRgb(gray, gray, gray);
                return true;
        }
    }

    public double MapLookupByteToComponent(int component, byte value) {
        IReadOnlyList<double>? ranges = _custom?.ComponentRanges;
        if (ranges == null || component < 0 || component * 2 + 1 >= ranges.Count) return value / 255D;
        double minimum = ranges[component * 2];
        return minimum + value / 255D * (ranges[component * 2 + 1] - minimum);
    }

    public static implicit operator PdfPageColorSpace(PdfPageColorSpaceKind kind) => new PdfPageColorSpace(kind);

    public static bool operator ==(PdfPageColorSpace left, PdfPageColorSpaceKind right) => left.Kind == right;
    public static bool operator !=(PdfPageColorSpace left, PdfPageColorSpaceKind right) => left.Kind != right;
    public static bool operator ==(PdfPageColorSpaceKind left, PdfPageColorSpace right) => left == right.Kind;
    public static bool operator !=(PdfPageColorSpaceKind left, PdfPageColorSpace right) => left != right.Kind;

    public override bool Equals(object? obj) => obj is PdfPageColorSpace other && Kind == other.Kind && ReferenceEquals(_calRgb, other._calRgb) && ReferenceEquals(_custom, other._custom);
    public override int GetHashCode() => (((int)Kind * 397) ^ (_calRgb?.GetHashCode() ?? 0)) * 397 ^ (_custom?.GetHashCode() ?? 0);

    private static int ComponentCountFor(PdfPageColorSpaceKind kind) => kind switch {
        PdfPageColorSpaceKind.DeviceRgb => 3,
        PdfPageColorSpaceKind.DeviceCmyk => 4,
        _ => 1
    };

    private static IReadOnlyList<double> NormalizeIccComponents(
        IReadOnlyList<double> components,
        int componentCount,
        IReadOnlyList<double>? ranges) {
        if (ranges == null) return components;
        var normalized = new double[componentCount];
        for (int index = 0; index < componentCount; index++) {
            double minimum = ranges[index * 2];
            double maximum = ranges[index * 2 + 1];
            double value = components[index];
            normalized[index] = value <= minimum ? 0D : value >= maximum ? 1D : (value - minimum) / (maximum - minimum);
        }
        return normalized;
    }

    private static IReadOnlyList<double> ClipIccComponents(
        IReadOnlyList<double> components,
        int componentCount,
        IReadOnlyList<double>? ranges) {
        if (ranges == null) return components;
        var clipped = new double[componentCount];
        for (int index = 0; index < componentCount; index++) {
            double minimum = ranges[index * 2];
            double maximum = ranges[index * 2 + 1];
            clipped[index] = Clamp(components[index], minimum, maximum);
        }
        return clipped;
    }

    private static byte ToByte(double value) => (byte)Math.Round(Clamp01(value) * 255D);
    private static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
    private static double Clamp(double value, double minimum, double maximum) => value < minimum ? minimum : value > maximum ? maximum : value;
    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private sealed class PdfPageCustomColorSpace {
        public PdfPageCustomColorSpace(int componentCount, bool usesIccApproximation) {
            ComponentCount = componentCount;
            UsesIccApproximation = usesIccApproximation;
        }

        public PdfPageCustomColorSpace(
            int componentCount,
            Func<IReadOnlyList<double>, OfficeIccRenderingIntent, OfficeColor?> colorTransform,
            bool usesIccApproximation = false,
            IReadOnlyList<double>? componentRanges = null) {
            ComponentCount = componentCount;
            ColorTransform = colorTransform;
            UsesIccApproximation = usesIccApproximation;
            ComponentRanges = componentRanges;
        }

        public PdfPageCustomColorSpace(IReadOnlyList<OfficeColor> palette, bool usesIccApproximation) {
            Palette = palette;
            ComponentCount = 1;
            UsesIccApproximation = usesIccApproximation;
        }

        public PdfPageCustomColorSpace(
            PdfPageColorSpace indexedBaseColorSpace,
            IReadOnlyList<IReadOnlyList<double>> indexedLookupComponents) {
            IndexedBaseColorSpace = indexedBaseColorSpace;
            IndexedLookupComponents = indexedLookupComponents;
            ComponentCount = 1;
            UsesIccApproximation = indexedBaseColorSpace.UsesIccApproximation;
        }

        public PdfPageCustomColorSpace(
            int componentCount,
            PdfPageColorSpace alternate,
            PdfColorSpaceTintTransform transform,
            int evaluationCost,
            Func<int, bool>? evaluationBudget) {
            ComponentCount = componentCount;
            Alternate = alternate;
            Transform = (components, output) =>
                (evaluationCost <= 0 || evaluationBudget == null || evaluationBudget(evaluationCost)) &&
                transform(components, output);
            UsesIccApproximation = alternate.UsesIccApproximation;
        }

        public PdfPageCustomColorSpace(PdfPageColorSpace patternBaseColorSpace) {
            ComponentCount = patternBaseColorSpace.ComponentCount;
            Alternate = patternBaseColorSpace;
            UsesIccApproximation = patternBaseColorSpace.UsesIccApproximation;
        }

        public int ComponentCount { get; }
        public bool UsesIccApproximation { get; }
        public IReadOnlyList<OfficeColor>? Palette { get; }
        public PdfPageColorSpace? IndexedBaseColorSpace { get; }
        public IReadOnlyList<IReadOnlyList<double>>? IndexedLookupComponents { get; }
        public PdfPageColorSpace? Alternate { get; }
        public PdfColorSpaceTintTransform? Transform { get; }
        public Func<IReadOnlyList<double>, OfficeIccRenderingIntent, OfficeColor?>? ColorTransform { get; }
        public IReadOnlyList<double>? ComponentRanges { get; }
    }

    private sealed class PdfPageCalRgbParameters {
        public static readonly PdfPageCalRgbParameters Default = new PdfPageCalRgbParameters(
            0.9505D, 1D, 1.089D, IdentityGamma, IdentityMatrix);

        public PdfPageCalRgbParameters(
            double whiteX,
            double whiteY,
            double whiteZ,
            IReadOnlyList<double>? gamma,
            IReadOnlyList<double>? matrix) {
            WhiteX = whiteX;
            WhiteY = whiteY;
            WhiteZ = whiteZ;
            Gamma = CopyOrDefault(gamma, IdentityGamma);
            Matrix = CopyOrDefault(matrix, IdentityMatrix);
        }

        public double WhiteX { get; }
        public double WhiteY { get; }
        public double WhiteZ { get; }
        public IReadOnlyList<double> Gamma { get; }
        public IReadOnlyList<double> Matrix { get; }

        private static System.Collections.ObjectModel.ReadOnlyCollection<double> CopyOrDefault(IReadOnlyList<double>? values, double[] fallback) {
            double[] copy = new double[fallback.Length];
            for (int i = 0; i < copy.Length; i++) copy[i] = values != null && i < values.Count ? values[i] : fallback[i];
            return Array.AsReadOnly(copy);
        }
    }
}
