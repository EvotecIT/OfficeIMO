using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageColorSpace {
    private static readonly double[] IdentityGamma = { 1D, 1D, 1D };
    private static readonly double[] IdentityMatrix = { 1D, 0D, 0D, 0D, 1D, 0D, 0D, 0D, 1D };
    private readonly PdfPageCalRgbParameters? _calRgb;

    public PdfPageColorSpace(PdfPageColorSpaceKind kind) {
        Kind = kind;
        _calRgb = null;
    }

    private PdfPageColorSpace(PdfPageCalRgbParameters calRgb) {
        Kind = PdfPageColorSpaceKind.CalRgb;
        _calRgb = calRgb;
    }

    public PdfPageColorSpaceKind Kind { get; }

    public static PdfPageColorSpace CalRgb(
        double whiteX,
        double whiteY,
        double whiteZ,
        IReadOnlyList<double>? gamma,
        IReadOnlyList<double>? matrix) =>
        new PdfPageColorSpace(new PdfPageCalRgbParameters(whiteX, whiteY, whiteZ, gamma, matrix));

    public OfficeColor ConvertCalRgb(double red, double green, double blue) {
        PdfPageCalRgbParameters parameters = _calRgb ?? PdfPageCalRgbParameters.Default;
        return OfficeColorSpaceConverter.FromCalibratedRgb(
            red, green, blue,
            parameters.WhiteX, parameters.WhiteY, parameters.WhiteZ,
            parameters.Gamma, parameters.Matrix);
    }

    public static implicit operator PdfPageColorSpace(PdfPageColorSpaceKind kind) => new PdfPageColorSpace(kind);

    public static bool operator ==(PdfPageColorSpace left, PdfPageColorSpaceKind right) => left.Kind == right;
    public static bool operator !=(PdfPageColorSpace left, PdfPageColorSpaceKind right) => left.Kind != right;
    public static bool operator ==(PdfPageColorSpaceKind left, PdfPageColorSpace right) => left == right.Kind;
    public static bool operator !=(PdfPageColorSpaceKind left, PdfPageColorSpace right) => left != right.Kind;

    public override bool Equals(object? obj) => obj is PdfPageColorSpace other && Kind == other.Kind && ReferenceEquals(_calRgb, other._calRgb);
    public override int GetHashCode() => ((int)Kind * 397) ^ (_calRgb?.GetHashCode() ?? 0);

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
