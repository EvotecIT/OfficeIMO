using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word;

public partial class WordShape {
    // Literal views are consumed by the native PDF adapter when no richer DrawingML fill is available.
    internal string FillColorValue => ReadLiteralShapeColor(stroke: false) ?? string.Empty;
    internal string? StrokeColorValue => ReadLiteralShapeColor(stroke: true);

    /// <summary>Gets the resolved fill color. Setting a role retains DrawingML theme slots and resolves VML colors against the document.</summary>
    public OfficeColor FillColor {
        get => ReadShapeColor(stroke: false);
        set => SetShapeColor(value, stroke: false);
    }

    /// <summary>Gets the resolved outline color. Setting a role retains DrawingML theme slots and resolves VML colors against the document.</summary>
    public OfficeColor StrokeColor {
        get => ReadShapeColor(stroke: true);
        set => SetShapeColor(value, stroke: true);
    }

    private OfficeColor ReadShapeColor(bool stroke) {
        var properties = _wpsShape?.GetFirstChild<Wps.ShapeProperties>();
        if (properties != null) {
            var fill = stroke
                ? properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()
                : properties.GetFirstChild<A.SolidFill>();
            var color = WordRunColor.ReadDrawingFill(fill, _document.ReadColorScheme);
            return color.HasValue ? WordRunColor.Resolve(color.Value, _document) : OfficeColor.Transparent;
        }
        string? value = stroke
            ? _rectangle?.StrokeColor?.Value ?? _roundRectangle?.StrokeColor?.Value ?? _ellipse?.StrokeColor?.Value
                ?? _polygon?.StrokeColor?.Value ?? _line?.StrokeColor?.Value ?? _shape?.StrokeColor?.Value
            : _rectangle?.FillColor?.Value ?? _roundRectangle?.FillColor?.Value ?? _ellipse?.FillColor?.Value
                ?? _polygon?.FillColor?.Value ?? _shape?.FillColor?.Value;
        return string.IsNullOrEmpty(value) ? OfficeColor.Transparent : OfficeColor.Parse(NormalizeHexNoHash(value));
    }

    private string? ReadLiteralShapeColor(bool stroke) {
        var properties = _wpsShape?.GetFirstChild<Wps.ShapeProperties>();
        if (properties != null) {
            var fill = stroke
                ? properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()
                : properties.GetFirstChild<A.SolidFill>();
            var color = WordRunColor.ReadDrawingFill(fill, _document.ReadColorScheme);
            return color.HasValue ? WordRunColor.Resolve(color.Value, _document).ToRgbHex() : null;
        }
        var literal = ReadShapeColor(stroke);
        return literal == OfficeColor.Transparent ? null : literal.ToRgbHex();
    }

    private void SetShapeColor(OfficeColor color, bool stroke) {
        var properties = _wpsShape?.GetFirstChild<Wps.ShapeProperties>();
        if (properties != null) {
            if (stroke) {
                var outline = properties.GetFirstChild<A.Outline>();
                if (outline == null) {
                    outline = new A.Outline();
                    properties.AddChild(outline, true);
                }
                WordRunColor.SetDrawingFill(outline, color, _document);
            } else {
                WordRunColor.SetDrawingFill(properties, color, _document);
            }
            return;
        }

        string value = "#" + WordRunColor.Resolve(color, _document).ToRgbHex();
        if (stroke) {
            if (_rectangle != null) _rectangle.StrokeColor = value;
            if (_roundRectangle != null) _roundRectangle.StrokeColor = value;
            if (_ellipse != null) _ellipse.StrokeColor = value;
            if (_polygon != null) _polygon.StrokeColor = value;
            if (_line != null) _line.StrokeColor = value;
            if (_shape != null) _shape.StrokeColor = value;
        } else {
            if (_rectangle != null) _rectangle.FillColor = value;
            if (_roundRectangle != null) _roundRectangle.FillColor = value;
            if (_ellipse != null) _ellipse.FillColor = value;
            if (_polygon != null) _polygon.FillColor = value;
            if (_shape != null) _shape.FillColor = value;
        }
    }
}
