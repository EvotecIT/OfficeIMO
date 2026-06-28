using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class ContentStreamBuilder {
    private readonly StringBuilder _sb;

    public ContentStreamBuilder(StringBuilder sb) {
        Guard.NotNull(sb, nameof(sb));
        _sb = sb;
    }

    public ContentStreamBuilder SaveState() {
        _sb.Append("q\n");
        return this;
    }

    public ContentStreamBuilder RestoreState() {
        _sb.Append("Q\n");
        return this;
    }

    public ContentStreamBuilder GraphicsState(string resourceName) {
        Guard.NotNullOrWhiteSpace(resourceName, nameof(resourceName));
        if (resourceName[0] == '/') {
            _sb.Append(resourceName);
        } else {
            _sb.Append('/').Append(resourceName);
        }

        _sb.Append(" gs\n");
        return this;
    }

    public ContentStreamBuilder FillColor(PdfColor color) {
        _sb.Append(F(color.R)).Append(' ').Append(F(color.G)).Append(' ').Append(F(color.B)).Append(" rg\n");
        return this;
    }

    public ContentStreamBuilder StrokeColor(PdfColor color) {
        _sb.Append(F(color.R)).Append(' ').Append(F(color.G)).Append(' ').Append(F(color.B)).Append(" RG\n");
        return this;
    }

    public ContentStreamBuilder LineWidth(double width) {
        _sb.Append(F(width)).Append(" w\n");
        return this;
    }

    public ContentStreamBuilder LineCap(int lineCap) {
        _sb.Append(lineCap.ToString(CultureInfo.InvariantCulture)).Append(" J\n");
        return this;
    }

    public ContentStreamBuilder LineJoin(int lineJoin) {
        _sb.Append(lineJoin.ToString(CultureInfo.InvariantCulture)).Append(" j\n");
        return this;
    }

    public ContentStreamBuilder StrokeDash(params double[] pattern) {
        Guard.NotNull(pattern, nameof(pattern));
        _sb.Append('[');
        for (int i = 0; i < pattern.Length; i++) {
            if (i > 0) {
                _sb.Append(' ');
            }

            _sb.Append(F(pattern[i]));
        }

        _sb.Append("] 0 d\n");
        return this;
    }

    public ContentStreamBuilder Rectangle(double x, double y, double width, double height) {
        _sb.Append(F(x)).Append(' ').Append(F(y)).Append(' ').Append(F(width)).Append(' ').Append(F(height)).Append(" re");
        return this;
    }

    public ContentStreamBuilder FillPath() {
        _sb.Append(" f\n");
        return this;
    }

    public ContentStreamBuilder FillStrokePath() {
        _sb.Append(" B\n");
        return this;
    }

    public ContentStreamBuilder StrokePath() {
        _sb.Append(" S\n");
        return this;
    }

    public ContentStreamBuilder MoveTo(double x, double y) {
        _sb.Append(F(x)).Append(' ').Append(F(y)).Append(" m");
        return this;
    }

    public ContentStreamBuilder PathSeparator() {
        _sb.Append('\n');
        return this;
    }

    public ContentStreamBuilder LineTo(double x, double y) {
        _sb.Append(' ').Append(F(x)).Append(' ').Append(F(y)).Append(" l");
        return this;
    }

    public ContentStreamBuilder CubicTo(double x1, double y1, double x2, double y2, double x3, double y3) {
        _sb.Append(' ')
            .Append(F(x1)).Append(' ').Append(F(y1)).Append(' ')
            .Append(F(x2)).Append(' ').Append(F(y2)).Append(' ')
            .Append(F(x3)).Append(' ').Append(F(y3)).Append(" c");
        return this;
    }

    public ContentStreamBuilder ClosePath() {
        _sb.Append(" h");
        return this;
    }

    public ContentStreamBuilder EndPath() {
        _sb.Append(" n\n");
        return this;
    }

    public ContentStreamBuilder ClipPath() {
        _sb.Append(" W");
        return this;
    }

    public ContentStreamBuilder TransformMatrix(double a, double b, double c, double d, double e, double f) {
        _sb.Append(F(a)).Append(' ')
            .Append(F(b)).Append(' ')
            .Append(F(c)).Append(' ')
            .Append(F(d)).Append(' ')
            .Append(F(e)).Append(' ')
            .Append(F(f)).Append(" cm\n");
        return this;
    }

    public ContentStreamBuilder TransformMatrix(OfficeTransform transform) =>
        TransformMatrix(transform.M11, transform.M12, transform.M21, transform.M22, transform.OffsetX, transform.OffsetY);

    public ContentStreamBuilder XObject(string resourceName) {
        Guard.NotNullOrWhiteSpace(resourceName, nameof(resourceName));
        if (resourceName[0] == '/') {
            _sb.Append(resourceName);
        } else {
            _sb.Append('/').Append(resourceName);
        }

        _sb.Append(" Do\n");
        return this;
    }

    public ContentStreamBuilder Shading(string resourceName) {
        Guard.NotNullOrWhiteSpace(resourceName, nameof(resourceName));
        if (resourceName[0] == '/') {
            _sb.Append(resourceName);
        } else {
            _sb.Append('/').Append(resourceName);
        }

        _sb.Append(" sh\n");
        return this;
    }

    public ContentStreamBuilder BeginText() {
        _sb.Append("BT\n");
        return this;
    }

    public ContentStreamBuilder EndText() {
        _sb.Append("ET\n");
        return this;
    }

    public ContentStreamBuilder Font(string resourceName, double size) {
        Guard.NotNullOrWhiteSpace(resourceName, nameof(resourceName));
        _sb.Append('/').Append(resourceName).Append(' ').Append(F(size)).Append(" Tf\n");
        return this;
    }

    public ContentStreamBuilder TextLeading(double leading) {
        _sb.Append(F(leading)).Append(" TL\n");
        return this;
    }

    public ContentStreamBuilder TextMatrix(double x, double y) {
        return TextMatrix(1, 0, 0, 1, x, y);
    }

    public ContentStreamBuilder TextMatrix(double a, double b, double c, double d, double e, double f) {
        _sb.Append(F(a)).Append(' ')
            .Append(F(b)).Append(' ')
            .Append(F(c)).Append(' ')
            .Append(F(d)).Append(' ')
            .Append(F(e)).Append(' ')
            .Append(F(f)).Append(" Tm\n");
        return this;
    }

    public ContentStreamBuilder MoveText(double x, double y) {
        _sb.Append(F(x)).Append(' ').Append(F(y)).Append(" Td\n");
        return this;
    }

    public ContentStreamBuilder NextTextLine() {
        _sb.Append("T*\n");
        return this;
    }

    public ContentStreamBuilder WordSpacing(double spacing) {
        _sb.Append(F(spacing)).Append(" Tw\n");
        return this;
    }

    public ContentStreamBuilder TextRise(double rise) {
        _sb.Append(F(rise)).Append(" Ts\n");
        return this;
    }

    public ContentStreamBuilder ShowHexText(string hexText) {
        Guard.NotNull(hexText, nameof(hexText));
        _sb.Append('<').Append(hexText).Append("> Tj\n");
        return this;
    }

    private static string F(double value) {
        if (Math.Abs(value) < 0.0005D) {
            value = 0D;
        }

        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }
}
