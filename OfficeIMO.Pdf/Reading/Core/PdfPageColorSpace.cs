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

    public bool HasPatternBaseColorSpace => Kind == PdfPageColorSpaceKind.Pattern && _custom?.Alternate != null;

    public static PdfPageColorSpace CalRgb(
        double whiteX,
        double whiteY,
        double whiteZ,
        IReadOnlyList<double>? gamma,
        IReadOnlyList<double>? matrix) =>
        new PdfPageColorSpace(new PdfPageCalRgbParameters(whiteX, whiteY, whiteZ, gamma, matrix));

    public static PdfPageColorSpace IccBased(PdfPageColorSpaceKind alternateKind) =>
        new PdfPageColorSpace(alternateKind, new PdfPageCustomColorSpace(ComponentCountFor(alternateKind), true));

    public static PdfPageColorSpace Indexed(IReadOnlyList<OfficeColor> palette, bool usesIccApproximation) =>
        new PdfPageColorSpace(PdfPageColorSpaceKind.Indexed, new PdfPageCustomColorSpace(palette, usesIccApproximation));

    public static PdfPageColorSpace Alternate(
        PdfPageColorSpaceKind kind,
        int componentCount,
        PdfPageColorSpace alternate,
        Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform) =>
        new PdfPageColorSpace(kind, new PdfPageCustomColorSpace(componentCount, alternate, transform));

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
        color = OfficeColor.Black;
        if (components == null || components.Count < ComponentCount || Kind == PdfPageColorSpaceKind.Pattern ||
            components.Take(ComponentCount).Any(value => !IsFinite(value))) return false;

        if (Kind == PdfPageColorSpaceKind.Indexed) {
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
            IReadOnlyList<double>? transformed = _custom.Transform(components);
            return transformed != null && alternate.TryConvertColor(transformed, out color);
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

    private static byte ToByte(double value) => (byte)Math.Round(Clamp01(value) * 255D);
    private static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private sealed class PdfPageCustomColorSpace {
        public PdfPageCustomColorSpace(int componentCount, bool usesIccApproximation) {
            ComponentCount = componentCount;
            UsesIccApproximation = usesIccApproximation;
        }

        public PdfPageCustomColorSpace(IReadOnlyList<OfficeColor> palette, bool usesIccApproximation) {
            Palette = palette;
            ComponentCount = 1;
            UsesIccApproximation = usesIccApproximation;
        }

        public PdfPageCustomColorSpace(
            int componentCount,
            PdfPageColorSpace alternate,
            Func<IReadOnlyList<double>, IReadOnlyList<double>?> transform) {
            ComponentCount = componentCount;
            Alternate = alternate;
            Transform = transform;
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
        public PdfPageColorSpace? Alternate { get; }
        public Func<IReadOnlyList<double>, IReadOnlyList<double>?>? Transform { get; }
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
