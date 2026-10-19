using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Projects canonical drawing elements to version 1 portable report drawing JSON.</summary>
/// <remarks>This is a bounded projection, not a general drawing serializer. Unsupported paint,
/// layout, fonts or resources reject the entire export; no partial drawing is returned.</remarks>
public static partial class OfficeDrawingPortableExporter {
    /// <summary>Exports an ordered drawing, returning no JSON when any source element cannot be preserved.</summary>
    /// <param name="drawing">Canonical drawing in points with a local top-left origin.</param>
    /// <param name="options">Explicit font family and resource budgets.</param>
    /// <param name="json">Complete version 1 JSON, or null on unsupported content or exceeded budgets.</param>
    /// <param name="diagnostics">A stable failure code and source location, or an empty collection on success.</param>
    public static bool TryExportJson(OfficeDrawing drawing, OfficeDrawingPortableOptions options,
        out string? json, out IReadOnlyList<OfficeConversionFidelityDiagnostic> diagnostics) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate();
        json = null;
        try {
            Dimension(drawing.Width, "width"); Dimension(drawing.Height, "height");
            var context = new ProjectionContext(options);
            List<Element> elements = context.Project(drawing, OfficeTransform.Identity, null, 0, "elements");
            ValidateOutputDepth(elements, 1, options.MaximumDepth);
            json = Serialize(drawing.Width, drawing.Height, elements, options.MaximumJsonCharacters);
            diagnostics = Array.Empty<OfficeConversionFidelityDiagnostic>();
            return true;
        } catch (ProjectionException error) {
            diagnostics = Array.AsReadOnly(new[] { new OfficeConversionFidelityDiagnostic(error.Code,
                error.Message, OfficeConversionLossKind.Failure, "OfficeIMO.Core.PortableDrawing", error.Location) });
            return false;
        }
    }

    private static void ValidateOutputDepth(List<Element> elements, int depth, int maximum) {
        if (elements.Count > 0 && depth > maximum) Fail("DRAWING_DEPTH_LIMIT", "Portable elements exceed the configured nesting limit.", "elements");
        foreach (Element element in elements) if (element.Children != null) ValidateOutputDepth(element.Children, depth + 1, maximum);
    }

    private sealed partial class ProjectionContext {
        internal readonly OfficeDrawingPortableOptions Options;
        private int _elements, _commands, _textCharacters;
        internal ProjectionContext(OfficeDrawingPortableOptions options) { Options = options; }

        internal List<Element> Project(OfficeDrawing drawing, OfficeTransform transform, OfficeFontFaceCollection? inheritedFonts,
            int depth, string location) {
            if (depth > Options.MaximumDepth) Fail("DRAWING_DEPTH_LIMIT", "Drawing nesting exceeds the configured limit.", location);
            if (drawing.TextShapingProvider != null || drawing.TextShapingLanguage != null)
                Fail("DRAWING_TEXT_SHAPING_UNSUPPORTED", "The portable profile does not preserve an external shaping provider or language override.", location);
            OfficeFontFaceCollection fonts = inheritedFonts?.Clone() ?? new OfficeFontFaceCollection();
            fonts.AddRange(drawing.Fonts);
            if (fonts.FontProgramProvider != null || fonts.FontVariationResolver != null)
                Fail("DRAWING_FONT_PROVIDER_UNSUPPORTED", "External font programs and variable-font resolvers are outside the portable profile.", location);
            var output = new List<Element>();
            for (int i = 0; i < drawing.Elements.Count; i++) {
                string path = location + "[" + i + "]";
                CountElement(path);
                OfficeDrawingElement item = drawing.Elements[i];
                if (item is OfficeDrawingEffectGroup effect) {
                    if (effect.Opacity != 1 || effect.BlendMode != OfficeBlendMode.Normal || effect.SoftMask != null)
                        Fail("DRAWING_EFFECT_UNSUPPORTED", "Opacity groups, blending and soft masks are outside the portable profile.", path);
                    ValidateTransform(effect.Transform, path);
                    OfficeTransform next = effect.Transform.Then(transform);
                    ValidateTransform(next, path);
                    output.AddRange(Project(effect.InnerDrawing, next, fonts, depth + 1, path + ".elements"));
                } else if (item is OfficeDrawingGroup group) {
                    if (group.ActualText != null || group.FrameTransform?.HasTransform == true)
                        Fail("DRAWING_GROUP_UNSUPPORTED", "Logical replacement text and transformed clip frames are outside the portable profile.", path);
                    Rect clip = Clip(group.ClipPath, group.X, group.Y, transform, path);
                    Coordinate(group.ContentOffsetX, path); Coordinate(group.ContentOffsetY, path);
                    var next = OfficeTransform.Translate(group.X + group.ContentOffsetX, group.Y + group.ContentOffsetY).Then(transform);
                    ValidateTransform(next, path);
                    output.Add(new Element { Kind = "group", Clip = clip,
                        Children = Project(group.InnerDrawing, next, fonts, depth + 1, path + ".elements") });
                } else if (item is OfficeDrawingShape shape) output.Add(ProjectShape(shape, transform, path));
                else if (item is OfficeDrawingText text) output.Add(ProjectText(text, transform, fonts, path));
                else Fail("DRAWING_ELEMENT_UNSUPPORTED", "The portable profile accepts paths, positioned text and rectangular clip groups only.", path);
            }
            return output;
        }

        internal void CountElement(string location) {
            if (++_elements > Options.MaximumElements) Fail("DRAWING_ELEMENT_LIMIT", "Drawing elements exceed the configured limit.", location);
        }
        internal void RequireElements(int count, string location) {
            if (count > Options.MaximumElements - _elements) Fail("DRAWING_ELEMENT_LIMIT", "Drawing elements exceed the configured limit.", location);
        }
        internal void RequireCommands(int count, string location) {
            if (count > Options.MaximumCommands - _commands) Fail("DRAWING_COMMAND_LIMIT", "Drawing commands exceed the configured limit.", location);
        }
        internal void CountCommands(int count, string location) {
            if (count > Options.MaximumCommands - _commands) Fail("DRAWING_COMMAND_LIMIT", "Drawing commands exceed the configured limit.", location);
            _commands += count;
        }
        internal void CountText(int count, string location) {
            if (count > Options.MaximumTextCharacters - _textCharacters) Fail("DRAWING_TEXT_LIMIT", "Drawing text exceeds the configured limit.", location);
            _textCharacters += count;
        }
    }

    private static void ValidateTransform(OfficeTransform transform, string location) {
        if (transform.M12 != 0 || transform.M21 != 0 || transform.M11 <= 0 || transform.M11 != transform.M22)
            Fail("DRAWING_TRANSFORM_UNSUPPORTED", "Only positive uniform scale and translation are supported.", location);
        Coordinate(transform.M11, location); Coordinate(transform.OffsetX, location); Coordinate(transform.OffsetY, location);
    }
    private static Rect Clip(OfficeClipPath clip, double x, double y, OfficeTransform transform, string location) {
        if (clip.Kind != OfficeClipPathKind.Rectangle) Fail("DRAWING_CLIP_UNSUPPORTED", "Only rectangular clip paths are supported.", location);
        Length(clip.Width, location); Length(clip.Height, location);
        OfficePoint start = Point(new OfficePoint(x, y), transform, location);
        return new Rect(start.X, start.Y, Length(clip.Width * transform.M11, location), Length(clip.Height * transform.M11, location));
    }
    private static OfficePoint Point(OfficePoint point, OfficeTransform transform, string location) {
        Coordinate(point.X, location); Coordinate(point.Y, location);
        OfficePoint result = transform.TransformPoint(point);
        Coordinate(result.X, location); Coordinate(result.Y, location); return result;
    }
    private static void Coordinate(double value, string location) {
        if (double.IsNaN(value) || double.IsInfinity(value) || Math.Abs(value) > 1000000)
            Fail("DRAWING_COORDINATE_LIMIT", "Drawing coordinates must be finite and within one million points.", location);
    }
    private static double Length(double value, string location) {
        Coordinate(value, location);
        if (value <= 0) Fail("DRAWING_DIMENSION_INVALID", "Drawing dimensions must be positive.", location);
        return value;
    }
    private static void Dimension(double value, string location) {
        Length(value, location);
        if (value > 14400) Fail("DRAWING_DIMENSION_LIMIT", "Drawing width and height cannot exceed 14,400 points.", location);
    }
    private static void Fail(string code, string message, string location) => throw new ProjectionException(code, message, location);
    private sealed class ProjectionException : Exception {
        internal ProjectionException(string code, string message, string location) : base(message) { Code = code; Location = location; }
        internal string Code { get; }
        internal string Location { get; }
    }
    private sealed class Element {
        internal string Kind = "path";
        internal List<OfficePathCommand>? Commands;
        internal string? Fill, Stroke, LineCap, LineJoin, Text, Font, Color;
        internal double StrokeWidth, X, Y, FontSize;
        internal string FillRule = "nonzero";
        internal Rect? Clip;
        internal List<Element>? Children;
    }
    private readonly struct Rect {
        internal Rect(double x, double y, double width, double height) { X = x; Y = y; Width = width; Height = height; }
        internal readonly double X, Y, Width, Height;
    }
}
