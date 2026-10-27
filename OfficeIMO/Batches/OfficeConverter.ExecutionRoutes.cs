namespace OfficeIMO;

public static partial class OfficeConverter {
    private static readonly HashSet<string> CheckpointRouteIds = new(StringComparer.OrdinalIgnoreCase) {
        "doc-pdf", "docx-pdf", "xlsx-pdf", "pptx-pdf", "html-pdf", "markdown-pdf", "pdf-docx", "pdf-pptx", "pdf-html", "txt-pdf"
    };

    internal static bool IsBuiltinRoute(string id) => BuiltinRouteIds.Contains(id);
    internal static IReadOnlyList<OfficeConversionExecutionRoute> ExecutionRoutes => Routes
        .Select(capability => new OfficeConversionExecutionRoute(capability, CheckpointRouteIds.Contains(capability.Id)
            || Registrations.TryGetValue(capability.Id, out IOfficeConversionRoute? route) && route is IOfficeConversionCheckpointOwner)).ToArray();

    internal static OfficeConversionExecutionRoute? FindExecutionRoute(string? id) => ExecutionRoutes
        .SingleOrDefault(route => string.Equals(route.Id, id, StringComparison.OrdinalIgnoreCase));

    internal static OfficeConversionExecutionRoute? FindExecutionRoute(string sourceExtension, string targetExtension,
        IReadOnlyList<OfficeConversionExecutionRoute> routes) {
        var matches = routes.Where(route => route.SourceExtensions.Contains(sourceExtension, StringComparer.OrdinalIgnoreCase)
            && string.Equals(route.TargetExtension, targetExtension, StringComparison.OrdinalIgnoreCase)).ToArray();
        string? preferred = OfficeConversionCapabilityCatalog.PreferredRouteId(sourceExtension, targetExtension);
        if (preferred != null && matches.FirstOrDefault(route => route.Id == preferred) is { } defaultRoute) return defaultRoute;
        return matches.Length == 1 ? matches[0] : null;
    }

    internal static bool SupportsDirectoryRoute(string id) => Registrations.TryGetValue(id, out IOfficeConversionRoute? route)
        && route is IOfficeConversionDirectoryOwner;

    internal static OfficeConversionExecutionInput CaptureDirectoryInput(string path, OfficeConversionOptions options) {
        string id = SelectRoute(options).Id;
        if (Registrations.TryGetValue(id, out IOfficeConversionRoute? route) && route is IOfficeConversionDirectoryOwner directory)
            return directory.CaptureDirectoryInput(path, options, new OfficeOperationLimits {
                MaximumInputBytes = options.MaximumInputBytes, MaximumOutputBytes = options.MaximumOutputBytes,
                MaximumXmlCharactersInPart = options.MaximumXmlCharactersInPart
            });
        throw new NotSupportedException("The selected conversion route does not admit directory packages.");
    }

    internal static string? CaptureRegisteredConfiguration(string id, OfficeConversionOptions options) =>
        Registrations.TryGetValue(id, out IOfficeConversionRoute? route) && route is IOfficeConversionCheckpointOwner owner
            ? owner.CaptureConfiguration(options) : null;
}
