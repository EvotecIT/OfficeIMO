using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Security;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows;

/// <summary>
/// Canonical package-neutral operation catalog projected from the detailed OfficeIMO capability owners.
/// </summary>
public static class OfficeOperationCapabilityCatalog {
    /// <summary>Stable catalog identifier.</summary>
    public const string Id = "OfficeIMO.Operations";

    /// <summary>Machine-readable catalog schema version.</summary>
    public const int SchemaVersion = 1;

    /// <summary>Current package-neutral operation catalog.</summary>
    public static OfficeOperationCatalog Current { get; } = new OfficeOperationCatalog(Id, SchemaVersion, CreateRows());

    /// <summary>All operation rows in deterministic order.</summary>
    public static IReadOnlyList<OfficeOperationCapability> All => Current.Capabilities;

    /// <summary>Gets rows owned by one package.</summary>
    public static IReadOnlyList<OfficeOperationCapability> FindByPackageId(string packageId) =>
        Current.FindByPackageId(packageId);

    /// <summary>Gets rows associated with one filename extension.</summary>
    public static IReadOnlyList<OfficeOperationCapability> FindByExtension(string extension) =>
        Current.FindByExtension(extension);

    /// <summary>Serializes the current catalog as deterministic JSON.</summary>
    public static string ToJson() => Current.ToJson();

    /// <summary>Formats the current catalog as a deterministic Markdown support matrix.</summary>
    public static string ToMarkdown() => Current.ToMarkdown();

    private static OfficeOperationCapability[] CreateRows() {
        var rows = new List<OfficeOperationCapability>();

        AddLegacyCatalog(rows, "OfficeIMO.Word", "Word", "WordCompatibilityCatalog.Current",
            WordCompatibilityCatalog.Current, WordFormatCatalog.All);
        AddLegacyCatalog(rows, "OfficeIMO.Excel", "Excel", "ExcelCompatibilityCatalog.Xls",
            ExcelCompatibilityCatalog.Xls, ExcelFormatCatalog.All);
        AddLegacyCatalog(rows, "OfficeIMO.Excel", "Excel", "ExcelCompatibilityCatalog.Xlsb",
            ExcelCompatibilityCatalog.Xlsb, ExcelFormatCatalog.All);
        AddLegacyCatalog(rows, "OfficeIMO.PowerPoint", "PowerPoint", "PowerPointCompatibilityCatalog.Current",
            PowerPointCompatibilityCatalog.Current, PowerPointFormatCatalog.All);
        AddConversionRoutes(rows);
        AddProtectionRows(rows);
        AddProvenanceRows(rows);

        return rows.ToArray();
    }

    private static void AddLegacyCatalog(
        ICollection<OfficeOperationCapability> rows,
        string packageId,
        string modernFormatId,
        string publicApi,
        OfficeCapabilityCatalog catalog,
        IReadOnlyList<OfficeFormatDescriptor> formats) {
        var extensionsByFormat = formats.ToDictionary(
            static format => format.Id,
            static format => new[] { format.Extension },
            StringComparer.Ordinal);
        string[] modernExtensions = formats
            .Where(static format => format.Generation == OfficeFormatGeneration.Modern)
            .Select(static format => format.Extension)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static extension => extension, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        foreach (OfficeCapability capability in catalog.Capabilities) {
            string[] legacyExtensions = extensionsByFormat.TryGetValue(capability.FormatId, out string[]? known)
                ? known
                : Array.Empty<string>();
            AddLegacyRow(rows, catalog, capability, packageId, publicApi, OfficeOperationKind.Read,
                capability.LegacyImport, capability.FormatId, null, legacyExtensions, "legacy-import");
            AddLegacyRow(rows, catalog, capability, packageId, publicApi, OfficeOperationKind.Create,
                capability.NewLegacyWrite, capability.FormatId, null, legacyExtensions, "new-legacy-write");
            AddLegacyRow(rows, catalog, capability, packageId, publicApi, OfficeOperationKind.Edit,
                capability.LegacyRoundTrip, capability.FormatId, null, legacyExtensions, "legacy-round-trip-edit");
            AddLegacyRow(rows, catalog, capability, packageId, publicApi, OfficeOperationKind.Preserve,
                capability.LegacyRoundTrip, capability.FormatId, null, legacyExtensions, "legacy-round-trip-preserve");
            AddLegacyRow(rows, catalog, capability, packageId, publicApi, OfficeOperationKind.Convert,
                capability.ModernToLegacy, modernFormatId, capability.FormatId, modernExtensions, "modern-to-legacy");
            AddLegacyRow(rows, catalog, capability, packageId, publicApi, OfficeOperationKind.Convert,
                capability.LegacyToModern, capability.FormatId, modernFormatId, legacyExtensions, "legacy-to-modern");
        }
    }

    private static void AddLegacyRow(
        ICollection<OfficeOperationCapability> rows,
        OfficeCapabilityCatalog catalog,
        OfficeCapability capability,
        string packageId,
        string publicApi,
        OfficeOperationKind operation,
        OfficeCapabilityCoverageState state,
        string formatId,
        string? targetFormatId,
        IEnumerable<string> extensions,
        string lane) {
        rows.Add(new OfficeOperationCapability(
            "legacy:" + catalog.Id + ":" + capability.Id + ":" + lane,
            packageId,
            formatId,
            capability.Id,
            operation,
            Map(state),
            publicApi,
            catalog.Id + "::" + capability.Id + "::" + lane,
            catalog.Id,
            extensions,
            targetFormatId,
            capability.Note));
    }

    private static void AddConversionRoutes(ICollection<OfficeOperationCapability> rows) {
        foreach (OfficeConversionCapability route in OfficeConversionCapabilityCatalog.All) {
            OfficeOperationKind operation = IsImageTarget(route.TargetExtension)
                ? OfficeOperationKind.Export
                : OfficeOperationKind.Convert;
            rows.Add(new OfficeOperationCapability(
                "conversion:" + route.Id,
                route.PackageId,
                route.Source,
                route.Id,
                operation,
                route.SupportLevel == OfficeConversionSupportLevel.Targeted
                    ? OfficeOperationSupportState.Partial
                    : OfficeOperationSupportState.Supported,
                route.Api,
                route.SupportEvidence,
                "OfficeConversionCapabilityCatalog",
                route.SourceExtensions,
                route.Target,
                route.KnownLimitations));
        }
    }

    private static void AddProtectionRows(ICollection<OfficeOperationCapability> rows) {
        foreach (OfficeProtectionCapability capability in OfficeProtectionCapabilityCatalog.Current.Capabilities) {
            string[] packages = ProtectionPackages(capability).ToArray();
            string[] extensions = FormatExtensions(capability.FormatId).ToArray();
            foreach (string packageId in packages) {
                AddProtectionRow(rows, capability, packageId, OfficeOperationKind.Inspect, capability.Inspect, extensions, "inspect");
                AddProtectionRow(rows, capability, packageId, OfficeOperationKind.Read, capability.Open, extensions, "open");
                AddProtectionRow(rows, capability, packageId, OfficeOperationKind.Create, capability.Create, extensions, "create");
                AddProtectionRow(rows, capability, packageId, OfficeOperationKind.Validate, capability.Validate, extensions, "validate");
                AddProtectionRow(rows, capability, packageId, OfficeOperationKind.Edit, capability.Mutate, extensions, "mutate");
                AddProtectionRow(rows, capability, packageId, OfficeOperationKind.Remove, capability.Remove, extensions, "remove");
            }
        }
    }

    private static void AddProtectionRow(
        ICollection<OfficeOperationCapability> rows,
        OfficeProtectionCapability capability,
        string packageId,
        OfficeOperationKind operation,
        OfficeProtectionCoverageState state,
        IEnumerable<string> extensions,
        string suffix) {
        rows.Add(new OfficeOperationCapability(
            "protection:" + capability.Id + ":" + PackageSlug(packageId) + ":" + suffix,
            packageId,
            capability.FormatId,
            capability.Id,
            operation,
            Map(state),
            capability.Api,
            "OfficeProtectionCapabilityCatalog.Current::" + capability.Id + "::" + suffix,
            OfficeProtectionCapabilityCatalog.Current.Id,
            extensions,
            limitation: capability.Limitation));
    }

    private static void AddProvenanceRows(ICollection<OfficeOperationCapability> rows) {
        foreach (OfficeProvenanceWorkflowCapability capability in OfficeProvenanceWorkflowCatalog.All) {
            AddProvenanceRow(rows, capability, OfficeOperationKind.Inspect, capability.CanInspect, "inspect");
            AddProvenanceRow(rows, capability, OfficeOperationKind.Validate, capability.CanAssess, "assess");
            AddProvenanceRow(rows, capability, OfficeOperationKind.Remove, capability.CanRemove, "remove");
        }
    }

    private static void AddProvenanceRow(
        ICollection<OfficeOperationCapability> rows,
        OfficeProvenanceWorkflowCapability capability,
        OfficeOperationKind operation,
        bool supported,
        string suffix) {
        rows.Add(new OfficeOperationCapability(
            "provenance:" + capability.Id + ":" + suffix,
            capability.OwnerPackage,
            capability.Label,
            capability.Id,
            operation,
            supported ? OfficeOperationSupportState.Supported : OfficeOperationSupportState.Unsupported,
            "OfficeProvenanceWorkflowCatalog",
            OfficeProvenanceWorkflowCatalog.Id + "::" + capability.Id + "::" + suffix,
            OfficeProvenanceWorkflowCatalog.Id,
            capability.Extensions,
            limitation: capability.Notes));
    }

    private static OfficeOperationSupportState Map(OfficeCapabilityCoverageState state) => state switch {
        OfficeCapabilityCoverageState.Native or OfficeCapabilityCoverageState.Equivalent => OfficeOperationSupportState.Supported,
        OfficeCapabilityCoverageState.Approximated or OfficeCapabilityCoverageState.Rasterized or
            OfficeCapabilityCoverageState.EmbeddedSource or OfficeCapabilityCoverageState.Dropped => OfficeOperationSupportState.Partial,
        OfficeCapabilityCoverageState.PreservedOpaque => OfficeOperationSupportState.Preserved,
        OfficeCapabilityCoverageState.Blocked => OfficeOperationSupportState.Rejected,
        OfficeCapabilityCoverageState.NotImplemented => OfficeOperationSupportState.Unsupported,
        OfficeCapabilityCoverageState.NotApplicable => OfficeOperationSupportState.NotApplicable,
        _ => throw new ArgumentOutOfRangeException(nameof(state))
    };

    private static OfficeOperationSupportState Map(OfficeProtectionCoverageState state) => state switch {
        OfficeProtectionCoverageState.Supported => OfficeOperationSupportState.Supported,
        OfficeProtectionCoverageState.Detected => OfficeOperationSupportState.Partial,
        OfficeProtectionCoverageState.Preserved => OfficeOperationSupportState.Preserved,
        OfficeProtectionCoverageState.Blocked => OfficeOperationSupportState.Rejected,
        OfficeProtectionCoverageState.NotSupported => OfficeOperationSupportState.Unsupported,
        OfficeProtectionCoverageState.NotApplicable => OfficeOperationSupportState.NotApplicable,
        _ => throw new ArgumentOutOfRangeException(nameof(state))
    };

    private static bool IsImageTarget(string extension) => extension.Equals(".png", StringComparison.OrdinalIgnoreCase)
        || extension.Equals(".jpg", StringComparison.OrdinalIgnoreCase)
        || extension.Equals(".jpeg", StringComparison.OrdinalIgnoreCase)
        || extension.Equals(".tif", StringComparison.OrdinalIgnoreCase)
        || extension.Equals(".tiff", StringComparison.OrdinalIgnoreCase)
        || extension.Equals(".webp", StringComparison.OrdinalIgnoreCase)
        || extension.Equals(".svg", StringComparison.OrdinalIgnoreCase);

    private static IEnumerable<string> ProtectionPackages(OfficeProtectionCapability capability) {
        if (capability.PackageId.StartsWith("format package", StringComparison.OrdinalIgnoreCase)) {
            foreach (string package in FormatOwners(capability.FormatId)) yield return package;
            yield break;
        }
        foreach (string part in capability.PackageId.Split(new[] { '/', '+' }, StringSplitOptions.RemoveEmptyEntries)) {
            string package = part.Trim();
            if (package.StartsWith("OfficeIMO.", StringComparison.Ordinal)) yield return package;
        }
    }

    private static IEnumerable<string> FormatOwners(string formatId) {
        string upper = formatId.ToUpperInvariant();
        if (upper.Contains("DOC")) yield return "OfficeIMO.Word";
        if (upper.Contains("XLS")) yield return "OfficeIMO.Excel";
        if (upper.Contains("PPT")) yield return "OfficeIMO.PowerPoint";
        if (upper.Contains("VISIO")) yield return "OfficeIMO.Visio";
    }

    private static IEnumerable<string> FormatExtensions(string formatId) {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["DOC"] = ".doc", ["DOCX"] = ".docx", ["DOCM"] = ".docm",
            ["XLS"] = ".xls", ["XLSX"] = ".xlsx", ["XLSM"] = ".xlsm", ["XLSB"] = ".xlsb",
            ["PPT"] = ".ppt", ["PPTX"] = ".pptx", ["PPTM"] = ".pptm",
            ["VISIO"] = ".vsdx", ["ODT"] = ".odt", ["ODS"] = ".ods", ["ODP"] = ".odp",
            ["EPUB"] = ".epub", ["PDF"] = ".pdf", ["ONE"] = ".one", ["PST"] = ".pst",
            ["RTF"] = ".rtf", ["EML"] = ".eml", ["MSG"] = ".msg", ["TNEF"] = ".dat"
        };
        foreach (string token in formatId.Split(new[] { '/', ' ', '(', ')', '-' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (map.TryGetValue(token, out string? extension)) yield return extension;
        }
    }

    private static string PackageSlug(string packageId) => packageId
        .Replace("OfficeIMO.", string.Empty)
        .Replace('.', '-')
        .ToLowerInvariant();
}
