namespace OfficeIMO.Pdf;

internal struct Matrix2D {
    public double A, B, C, D, E, F;
    public Matrix2D(double a, double b, double c, double d, double e, double f) { A = a; B = b; C = c; D = d; E = e; F = f; }
    public static Matrix2D Identity => new Matrix2D(1, 0, 0, 1, 0, 0);
    public (double X, double Y) Transform(double x, double y) {
        double nx = A * x + C * y + E;
        double ny = B * x + D * y + F;
        return (nx, ny);
    }
    public static Matrix2D Multiply(Matrix2D m1, Matrix2D m2) {
        // standard affine multiply: m1 * m2
        return new Matrix2D(
            m1.A * m2.A + m1.B * m2.C,
            m1.A * m2.B + m1.B * m2.D,
            m1.C * m2.A + m1.D * m2.C,
            m1.C * m2.B + m1.D * m2.D,
            m1.E * m2.A + m1.F * m2.C + m2.E,
            m1.E * m2.B + m1.F * m2.D + m2.F);
    }
}

