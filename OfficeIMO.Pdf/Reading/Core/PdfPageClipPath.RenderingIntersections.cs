using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

// Clip intersections that geometry cannot resolve stay exact for rendering: the primary geometry
// (rectangle or Commands) is a superset, and renderers also clip by each of these paths.
// Every path enters through IntersectPathWithPath, which charges its flattened vertices to the
// page's clipping budget, so a chain of any depth stays exact and its total size stays bounded.
// Clips share chains, so extending one costs only the new paths.
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

    private PdfPageClipPath WithRenderingIntersections(IntersectionChain? intersections) {
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

    // A new clip usually extends the active chain, which costs only the new paths. Merging two chains
    // walks both, so that work is charged to the page clipping budget.
    private static IntersectionChain? CombineRenderingIntersections(
        IntersectionChain? active,
        IntersectionChain? clip,
        PdfPageClipPath[]? added,
        PdfTextClippingBudget? textClippingBudget) {
        IntersectionChain? combined = active;
        if (clip != null && !ReferenceEquals(clip, active)) {
            if (active == null) {
                combined = clip;
            } else {
                IntersectionChain longer = active.Count >= clip.Count ? active : clip;
                IntersectionChain shorter = ReferenceEquals(longer, active) ? clip : active;
                textClippingBudget?.ChargeLinearIntersectionWork((long)longer.Count + shorter.Count);
                combined = longer.AppendDistinct(shorter);
            }
        }

        if (added == null) return combined;
        for (int i = 0; i < added.Length; i++) {
            PdfPageClipPath path = added[i];
            if (i > 0 && !path.IsRectangle && ReferenceEquals(path.Commands, added[i - 1].Commands)) continue;
            combined = new IntersectionChain(path, combined);
        }
        return combined;
    }

    // An immutable list of rendering intersections, newest first. Clips that extend a chain share it.
    internal sealed class IntersectionChain {
        internal IntersectionChain(PdfPageClipPath path, IntersectionChain? previous) {
            Path = path;
            Previous = previous;
            Count = (previous?.Count ?? 0) + 1;
        }

        internal PdfPageClipPath Path { get; }

        internal IntersectionChain? Previous { get; }

        internal int Count { get; }

        /// <summary>The paths, oldest first.</summary>
        internal PdfPageClipPath[] ToArray() {
            var paths = new PdfPageClipPath[Count];
            IntersectionChain? node = this;
            for (int i = paths.Length - 1; i >= 0; i--) {
                paths[i] = node!.Path;
                node = node.Previous;
            }
            return paths;
        }

        // Paths sharing a command list are the same page-space clip, because Translate and WithBounds build new lists.
        internal IntersectionChain AppendDistinct(IntersectionChain other) {
            var seen = new HashSet<IReadOnlyList<OfficePathCommand>>(CommandListReferenceComparer.Instance);
            for (IntersectionChain? node = this; node != null; node = node.Previous) {
                if (ReferenceEquals(node, other)) return this;
                if (!node.Path.IsRectangle) seen.Add(node.Path.Commands);
            }

            IntersectionChain combined = this;
            PdfPageClipPath[] paths = other.ToArray();
            for (int i = 0; i < paths.Length; i++) {
                if (paths[i].IsRectangle || seen.Add(paths[i].Commands)) combined = new IntersectionChain(paths[i], combined);
            }
            return combined;
        }
    }

    private sealed class CommandListReferenceComparer : IEqualityComparer<IReadOnlyList<OfficePathCommand>> {
        internal static readonly CommandListReferenceComparer Instance = new CommandListReferenceComparer();
        public bool Equals(IReadOnlyList<OfficePathCommand>? x, IReadOnlyList<OfficePathCommand>? y) => ReferenceEquals(x, y);
        public int GetHashCode(IReadOnlyList<OfficePathCommand> obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }

    private OfficeClipPath AddOfficeClipIntersections(OfficeClipPath primary, double primitiveX, double primitiveY) {
        if (_renderingIntersections == null) return primary;
        PdfPageClipPath[] chain = _renderingIntersections.ToArray();
        var intersections = new List<OfficeClipPath>(chain.Length);
        try {
            for (int i = 0; i < chain.Length; i++) {
                PdfPageClipPath intersection = chain[i];
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
