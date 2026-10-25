namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Reuses pure container geometry only within this style computation and the complete numeric resolution context.</summary>
    private sealed class ContainerGeometryReuseCache {
        private const int MaximumEntries = 256;
        private readonly Dictionary<ContainerGeometryKey, (double Width, double? Height)> _geometry = new();

        internal (double Width, double? Height) Resolve(HtmlComputedStyle style,
            double containingWidth, double? containingHeight, double fontSize, double rootFontSize,
            MediaEnvironment environment, double containerUnitWidth, double containerUnitHeight) {
            var key = new ContainerGeometryKey(style.DeclarationIdentity, containingWidth, containingHeight,
                fontSize, rootFontSize, environment.Width, environment.Height, containerUnitWidth, containerUnitHeight);
            if (_geometry.TryGetValue(key, out var found)) return found;
            double width = ResolveContainerElementWidth(style, containingWidth, fontSize, rootFontSize,
                environment, containerUnitWidth, containerUnitHeight);
            double? height = ResolveContainerElementHeight(style, width, containingWidth, containingHeight,
                fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight);
            var result = (width, height);
            if (_geometry.Count < MaximumEntries) _geometry.Add(key, result);
            return result;
        }
    }

    private readonly record struct ContainerGeometryKey(object DeclarationIdentity,
        double ContainingWidth, double? ContainingHeight, double FontSize, double RootFontSize,
        double ViewportWidth, double ViewportHeight, double ContainerUnitWidth, double ContainerUnitHeight);
}
