using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

// Clip intersections that geometry cannot resolve stay exact for rendering: the primary geometry
// (rectangle or Commands) is a superset, and renderers also clip by each of these paths.
// Every path enters through IntersectPathWithPath, which charges its flattened vertices to the
// page's clipping budget, so a chain of any depth stays exact and its total size stays bounded.
internal readonly partial struct PdfPageClipPath {
    private PdfPageClipPath WithoutRenderingIntersections() {
        if (_renderingIntersections == null) return this;
        if (_rectangularPrimary) {
            return Rectangle(X, Y, Width, Height)
                .WithExactness(_isPrimaryExact)
                .WithTextClipping(ContainsTextClipping);
        }

        return new PdfPageClipPath(
            X,
            Y,
            Width,
            Height,
            IsRectangle,
            FillRule,
            Commands,
            _isPrimaryExact,
            ContainsTextClipping,
            _canServeAsExactPathClip);
    }

    private PdfPageClipPath WithRenderingIntersections(PdfPageClipPath[]? intersections) {
        if (intersections == null || Width <= 0D || Height <= 0D) {
            return WithoutRenderingIntersections();
        }

        bool isExact = _isPrimaryExact;
        if (!IsRectangle) {
            return new PdfPageClipPath(
                X,
                Y,
                Width,
                Height,
                false,
                FillRule,
                Commands,
                isExact,
                ContainsTextClipping,
                _canServeAsExactPathClip,
                intersections,
                _rectangularPrimary);
        }

        return new PdfPageClipPath(
            X,
            Y,
            Width,
            Height,
            false,
            OfficeFillRule.NonZero,
            CreateRectangleCommands(X, Y, Width, Height),
            isExact,
            ContainsTextClipping,
            isExact,
            intersections,
            rectangularPrimary: true);
    }

    private static PdfPageClipPath[]? CombineRenderingIntersections(
        PdfPageClipPath[]? first,
        PdfPageClipPath[]? second,
        PdfPageClipPath[]? third) {
        int count = (first?.Length ?? 0) + (second?.Length ?? 0) + (third?.Length ?? 0);
        if (count == 0) return null;
        if (second == null && third == null) return first;
        if (first == null && third == null) return second;
        if (first == null && second == null) return third;
        if (ReferenceEquals(first, second) && third == null) return first;

        var combined = new List<PdfPageClipPath>(count);
        var seen = new HashSet<IReadOnlyList<OfficePathCommand>>(CommandListReferenceComparer.Instance);
        AddDistinctRenderingIntersections(combined, seen, first);
        AddDistinctRenderingIntersections(combined, seen, second);
        AddDistinctRenderingIntersections(combined, seen, third);
        return combined.ToArray();
    }

    // Linear in the chain length. Paths sharing a command list are the same page-space clip,
    // because Translate and WithBounds build new command lists.
    private static void AddDistinctRenderingIntersections(
        List<PdfPageClipPath> combined,
        HashSet<IReadOnlyList<OfficePathCommand>> seen,
        PdfPageClipPath[]? intersections) {
        if (intersections == null) return;
        for (int i = 0; i < intersections.Length; i++) {
            PdfPageClipPath intersection = intersections[i];
            if (intersection.IsRectangle || seen.Add(intersection.Commands)) combined.Add(intersection);
        }
    }

    private sealed class CommandListReferenceComparer : IEqualityComparer<IReadOnlyList<OfficePathCommand>> {
        internal static readonly CommandListReferenceComparer Instance = new CommandListReferenceComparer();
        public bool Equals(IReadOnlyList<OfficePathCommand>? x, IReadOnlyList<OfficePathCommand>? y) => ReferenceEquals(x, y);
        public int GetHashCode(IReadOnlyList<OfficePathCommand> obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }

    private OfficeClipPath AddOfficeClipIntersections(OfficeClipPath primary, double primitiveX, double primitiveY) {
        if (_renderingIntersections == null) return primary;
        var intersections = new List<OfficeClipPath>(_renderingIntersections.Length);
        try {
            for (int i = 0; i < _renderingIntersections.Length; i++) {
                PdfPageClipPath intersection = _renderingIntersections[i];
                IReadOnlyList<OfficePathCommand> commands = intersection.IsRectangle
                    ? CreateRectangleCommands(intersection.X - primitiveX, intersection.Y - primitiveY, intersection.Width, intersection.Height)
                    : intersection.TranslateCommands(primitiveX, primitiveY);
                // Keep the intersection's offsets in the primary clip's local coordinates.
                intersections.Add(OfficeClipPath.Path(Width, Height, commands, intersection.FillRule));
            }
        } catch (ArgumentException) {
            // The primary clip alone is a superset; never drop the clip entirely.
            return primary;
        }

        return primary.WithIntersections(intersections);
    }

    private static List<OfficePathCommand> CreateRectangleCommands(double x, double y, double width, double height) {
        double right = x + width;
        double bottom = y + height;
        return new List<OfficePathCommand>(5) {
            OfficePathCommand.MoveTo(new OfficePoint(x, y)),
            OfficePathCommand.LineTo(new OfficePoint(right, y)),
            OfficePathCommand.LineTo(new OfficePoint(right, bottom)),
            OfficePathCommand.LineTo(new OfficePoint(x, bottom)),
            OfficePathCommand.Close()
        };
    }
}
