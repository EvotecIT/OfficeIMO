using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Composes native cached reflections and modeled transforms for rendering and page-coordinate diagram operations.</summary>
internal readonly struct VisioNativeShapeTransform {
    private VisioNativeShapeTransform(OfficeTransform matrix, bool flipX, bool flipY,
        bool hasReflection, double textSign, double textRotation) {
        Matrix = matrix; FlipX = flipX; FlipY = flipY; HasReflection = hasReflection;
        _textSign = textSign; _textRotation = textRotation;
    }

    internal OfficeTransform Matrix { get; }
    internal bool FlipX { get; }
    internal bool FlipY { get; }
    internal bool HasReflection { get; }
    private readonly double _textSign, _textRotation;

    internal OfficePoint PagePoint(double x, double y) => Matrix.TransformPoint(new OfficePoint(x, y));

    internal OfficePoint LocalPoint(OfficePoint pagePoint) => Matrix.Invert().TransformPoint(pagePoint);

    internal static void MoveInPage(VisioShape shape, double deltaX, double deltaY) {
        OfficePoint pin = new(shape.PinX, shape.PinY);
        if (shape.Parent != null) {
            VisioNativeShapeTransform parent = Create(shape.Parent);
            OfficePoint pagePin = parent.PagePoint(pin.X, pin.Y);
            pin = parent.LocalPoint(new OfficePoint(pagePin.X + deltaX, pagePin.Y + deltaY));
        } else pin = new OfficePoint(pin.X + deltaX, pin.Y + deltaY);
        shape.PinX = pin.X; shape.PinY = pin.Y;
    }

    // Text remains readable when a shape or containing group is reflected. Each odd
    // reflection reverses the preceding text angle before the containing rotation.
    internal double TextAngle(double angle) => angle * _textSign + _textRotation;

    internal static VisioNativeShapeTransform Create(VisioShape shape,
        ICollection<OfficeImageExportDiagnostic>? diagnostics = null, string? source = null) {
        OfficeTransform matrix = OfficeTransform.Identity;
        bool ownX = false, ownY = false, reflected = false;
        double textSign = 1, textRotation = 0;
        for (VisioShape? current = shape; current != null; current = current.Parent) {
            bool flipX = ReadReflection(current, "FlipX", diagnostics, source);
            bool flipY = ReadReflection(current, "FlipY", diagnostics, source);
            if (ReferenceEquals(current, shape)) { ownX = flipX; ownY = flipY; }
            reflected |= flipX || flipY;
            if (flipX != flipY) { textSign = -textSign; textRotation = -textRotation; }
            textRotation += current.Angle;
            matrix = matrix.Then(OfficeTransform.Translate(-current.LocPinX, -current.LocPinY))
                .Then(OfficeTransform.Scale(flipX ? -1 : 1, flipY ? -1 : 1))
                .Then(OfficeTransform.RotateDegrees(OfficeGeometry.RadiansToDegrees(current.Angle)))
                .Then(OfficeTransform.Translate(current.PinX, current.PinY));
        }
        return new VisioNativeShapeTransform(matrix, ownX, ownY, reflected, textSign, textRotation);
    }

    private static bool ReadReflection(VisioShape shape, string name,
        ICollection<OfficeImageExportDiagnostic>? diagnostics, string? source) {
        XElement? Find(VisioShape owner) => owner.PreservedCellElements.FirstOrDefault(cell => (string?)cell.Attribute("N") == name);
        XElement? cell = Find(shape);
        if (cell == null || (string?)cell.Attribute("F") == "Inh")
            cell = (shape.MasterShape ?? shape.Master?.Shape) is VisioShape master ? Find(master) ?? cell : cell;
        if (cell == null) return false;
        string? formula = (string?)cell.Attribute("F");
        bool literalFormula = VisioShapeGeometry.TryParseLiteralWithoutShape(formula, out double literal);
        if (!string.IsNullOrWhiteSpace(formula) && formula != "Inh" && !literalFormula)
            diagnostics?.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                "VISIO_SHAPE_CACHED_REFLECTION", "Shape rendering uses the cached value of an unsupported " + name + " formula.",
                (source ?? "Visio page") + " / " + (shape.NameU ?? shape.Id), OfficeConversionLossKind.Approximation));
        if (VisioShapeGeometry.TryParseLiteralWithoutShape((string?)cell.Attribute("V"), out double value)) return value != 0;
        if (literalFormula) return literal != 0;
        throw new InvalidDataException("Reflection cell " + name + " has no usable finite value.");
    }

    internal static void ReportInvalid(VisioShape shape, ICollection<OfficeImageExportDiagnostic>? diagnostics,
        string? source, Exception exception) => diagnostics?.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning, "VISIO_SHAPE_TRANSFORM_INVALID", exception.Message,
            (source ?? "Visio page") + " / " + (shape.NameU ?? shape.Id), OfficeConversionLossKind.Omission));

    internal static void ReportArtwork(VisioShape shape, ICollection<OfficeImageExportDiagnostic>? diagnostics, string? source) =>
        diagnostics?.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
            "VISIO_SHAPE_REFLECTION_ARTWORK", "Decorative stencil artwork or a database fallback does not reproduce native reflection.",
            (source ?? "Visio page") + " / " + (shape.NameU ?? shape.Id), OfficeConversionLossKind.Approximation));
}
