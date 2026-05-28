namespace OfficeIMO.Pdf;

internal struct Matrix2D {
    public double A, B, C, D, E, F;
    public Matrix2D(double a, double b, double c, double d, double e, double f) { A = a; B = b; C = c; D = d; E = e; F = f; }
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
}

