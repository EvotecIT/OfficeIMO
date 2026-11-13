using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Convert;

/// <summary>The routes `officeimo convert` executes: the shared workflow routes plus the structured-data routes that <see cref="OfficeConverter"/> runs itself.</summary>
internal static class ConvertRoutes {
    internal static IReadOnlyList<OfficeConversionCapability> DataRoutes { get; } = OfficeConverter.Routes
        .Where(route => route.PackageId is "OfficeIMO.Data.Conversions" or "OfficeIMO.Data.Conversions.Excel")
        .ToArray();

    internal static OfficeConversionCapability? FindData(string sourceExtension, string targetExtension) {
        OfficeConversionCapability[] matches = DataRoutes.Where(route =>
            route.SourceExtensions.Contains(sourceExtension, StringComparer.OrdinalIgnoreCase)
            && string.Equals(route.TargetExtension, targetExtension, StringComparison.OrdinalIgnoreCase)).ToArray();
        return matches.Length == 1 ? matches[0] : null;
    }

    internal static bool IsTarget(string targetExtension) =>
        OfficeWorkflowCatalog.ExecutableRoutes.Any(route => route.TargetExtension == targetExtension)
        || DataRoutes.Any(route => route.TargetExtension == targetExtension);
}
