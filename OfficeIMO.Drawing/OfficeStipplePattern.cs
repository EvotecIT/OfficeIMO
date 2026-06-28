namespace OfficeIMO.Drawing;

internal readonly struct OfficeStipplePattern {
    private const int TileSize = 4;
    private readonly OfficeHatchPatternKind _kind;

    private OfficeStipplePattern(OfficeHatchPatternKind kind) {
        _kind = kind;
    }

    internal static bool TryCreate(OfficeHatchPatternKind kind, out OfficeStipplePattern pattern) {
        switch (kind) {
            case OfficeHatchPatternKind.Percent6_25:
            case OfficeHatchPatternKind.Percent12_5:
            case OfficeHatchPatternKind.Percent25:
            case OfficeHatchPatternKind.Percent50:
            case OfficeHatchPatternKind.Percent75:
                pattern = new OfficeStipplePattern(kind);
                return true;
            default:
                pattern = default;
                return false;
        }
    }

    internal int Size => TileSize;

    internal bool IsFilled(int x, int y) {
        x = PositiveModulo(x, TileSize);
        y = PositiveModulo(y, TileSize);

        switch (_kind) {
            case OfficeHatchPatternKind.Percent6_25:
                return x == 0 && y == 0;
            case OfficeHatchPatternKind.Percent12_5:
                return (x == 0 && y == 0) || (x == 2 && y == 2);
            case OfficeHatchPatternKind.Percent25:
                return (x == 0 || x == 2) && (y == 0 || y == 2);
            case OfficeHatchPatternKind.Percent50:
                return ((x + y) % 2) == 0;
            case OfficeHatchPatternKind.Percent75:
                return !((x == 1 && y == 0) || (x == 3 && y == 1) || (x == 0 && y == 2) || (x == 2 && y == 3));
            default:
                return false;
        }
    }

    private static int PositiveModulo(int value, int divisor) {
        int remainder = value % divisor;
        return remainder < 0 ? remainder + divisor : remainder;
    }
}
