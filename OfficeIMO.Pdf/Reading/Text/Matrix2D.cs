namespace OfficeIMO.Pdf;

internal struct Matrix2D : IEquatable<Matrix2D> {
    public double A, B, C, D, E, F;
    public Matrix2D(double a, double b, double c, double d, double e, double f) { A = a; B = b; C = c; D = d; E = e; F = f; }
    // Matrices key the per-page caches of form and glyph analysis. The default struct hash uses only A, so placements
    // that differ by translation shared one bucket; equality stays component-wise double equality.
    public readonly bool Equals(Matrix2D other) =>
        A.Equals(other.A) && B.Equals(other.B) && C.Equals(other.C) &&
        D.Equals(other.D) && E.Equals(other.E) && F.Equals(other.F);

    public override readonly bool Equals(object? obj) => obj is Matrix2D other && Equals(other);

    public override readonly int GetHashCode() {
        unchecked {
            int hash = HashComponent(A);
            hash = (hash * 31) + HashComponent(B);
            hash = (hash * 31) + HashComponent(C);
            hash = (hash * 31) + HashComponent(D);
            hash = (hash * 31) + HashComponent(E);
            return (hash * 31) + HashComponent(F);
        }
    }

    // Zero of either sign, and every NaN, are equal doubles and must hash alike.
    internal static int HashComponent(double value) => double.IsNaN(value) ? 1 : value == 0D ? 0 : value.GetHashCode();
    public static Matrix2D Identity => new Matrix2D(1, 0, 0, 1, 0, 0);
    public static Matrix2D Translation(double x, double y) => new Matrix2D(1, 0, 0, 1, x, y);
    public (double X, double Y) Transform(double x, double y) {
        double nx = A * x + C * y + E;
        double ny = B * x + D * y + F;
        return (nx, ny);
    }
    public static Matrix2D Multiply(Matrix2D m1, Matrix2D m2) {
        // Column-vector affine multiply: apply m2 in the coordinate system already transformed by m1.
        return new Matrix2D(
            m1.A * m2.A + m1.C * m2.B,
            m1.B * m2.A + m1.D * m2.B,
            m1.A * m2.C + m1.C * m2.D,
            m1.B * m2.C + m1.D * m2.D,
            m1.A * m2.E + m1.C * m2.F + m1.E,
            m1.B * m2.E + m1.D * m2.F + m1.F);
    }

    internal bool IsConformalStrokeTransform() {
        double firstLengthSquared = (A * A) + (B * B);
        double secondLengthSquared = (C * C) + (D * D);
        if (firstLengthSquared <= 0D || secondLengthSquared <= 0D ||
            double.IsNaN(firstLengthSquared) || double.IsNaN(secondLengthSquared) ||
            double.IsInfinity(firstLengthSquared) || double.IsInfinity(secondLengthSquared)) return false;

        double dot = (A * C) + (B * D);
        return firstLengthSquared == secondLengthSquared && dot == 0D;
    }
}
