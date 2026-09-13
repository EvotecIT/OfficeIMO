using System.Collections.ObjectModel;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows;

/// <summary>One exact file-format registration in a provenance workflow capability.</summary>
public sealed class OfficeProvenanceWorkflowFormat {
    internal OfficeProvenanceWorkflowFormat(
        string extension,
        IEnumerable<OfficeProvenanceAssetFormat> assetFormats,
        bool memoryOnlyAvailable = false,
        bool browserAvailable = false) {
        if (string.IsNullOrWhiteSpace(extension) || extension[0] != '.') {
            throw new ArgumentException("A provenance format extension must begin with '.'.", nameof(extension));
        }
        OfficeProvenanceAssetFormat[] registeredAssetFormats = assetFormats?.Distinct().OrderBy(static item => item).ToArray()
            ?? throw new ArgumentNullException(nameof(assetFormats));
        if (registeredAssetFormats.Length == 0 || registeredAssetFormats.Contains(OfficeProvenanceAssetFormat.Unknown)) {
            throw new ArgumentException("A provenance format requires at least one known structural format.", nameof(assetFormats));
        }
        if (browserAvailable && !memoryOnlyAvailable) {
            throw new ArgumentException("A browser format must also support the memory-only workflow.", nameof(browserAvailable));
        }

        Extension = extension.ToLowerInvariant();
        AssetFormats = Array.AsReadOnly(registeredAssetFormats);
        MemoryOnlyAvailable = memoryOnlyAvailable;
        BrowserAvailable = browserAvailable;
    }

    /// <summary>Registered filename extension.</summary>
    public string Extension { get; }
    /// <summary>Structural formats that the file contents may match.</summary>
    public IReadOnlyList<OfficeProvenanceAssetFormat> AssetFormats { get; }
    /// <summary>Whether the shared byte-oriented workflow is qualified for this extension.</summary>
    public bool MemoryOnlyAvailable { get; }
    /// <summary>Whether the static browser host is qualified for this extension.</summary>
    public bool BrowserAvailable { get; }
}

/// <summary>One discoverable provenance workflow capability owned by an OfficeIMO package.</summary>
public sealed class OfficeProvenanceWorkflowCapability {
    private readonly IReadOnlyList<string> _extensions;
    private readonly IReadOnlyList<string> _memoryOnlyExtensions;
    private readonly bool _browserAvailable;

    internal OfficeProvenanceWorkflowCapability(
        string id,
        string label,
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner owner,
        IEnumerable<OfficeProvenanceWorkflowFormat> formats,
        bool canInspect,
        bool canAssess,
        bool canRemove,
        string notes,
        string? browserLabel = null,
        int browserOrder = int.MaxValue) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Capability id cannot be empty.", nameof(id));
        if (string.IsNullOrWhiteSpace(label)) throw new ArgumentException("Capability label cannot be empty.", nameof(label));
        OfficeProvenanceWorkflowFormat[] registeredFormats = formats?.OrderBy(static item => item.Extension, StringComparer.Ordinal).ToArray()
            ?? throw new ArgumentNullException(nameof(formats));
        if (registeredFormats.Length == 0) {
            throw new ArgumentException("A provenance capability must register at least one OfficeIMO-owned format.", nameof(formats));
        }
        string[] duplicateExtensions = registeredFormats
            .GroupBy(static item => item.Extension, StringComparer.OrdinalIgnoreCase)
            .Where(static group => group.Count() > 1)
            .Select(static group => group.Key)
            .ToArray();
        if (duplicateExtensions.Length != 0) {
            throw new ArgumentException("A provenance capability cannot register duplicate extensions: " + string.Join(", ", duplicateExtensions), nameof(formats));
        }
        bool hasBrowserFormats = registeredFormats.Any(static item => item.BrowserAvailable);
        if (hasBrowserFormats != !string.IsNullOrWhiteSpace(browserLabel)) {
            throw new ArgumentException("Browser-qualified capabilities require one browser label, and other capabilities must omit it.", nameof(browserLabel));
        }

        Id = id.Trim();
        Label = label.Trim();
        Owner = owner;
        OwnerPackage = OfficeProvenanceWorkflowCatalog.GetOwnerPackage(owner);
        Formats = Array.AsReadOnly(registeredFormats);
        _extensions = Array.AsReadOnly(registeredFormats.Select(static item => item.Extension).ToArray());
        _memoryOnlyExtensions = Array.AsReadOnly(registeredFormats
            .Where(static item => item.MemoryOnlyAvailable)
            .Select(static item => item.Extension)
            .ToArray());
        _browserAvailable = hasBrowserFormats;
        CanInspect = canInspect;
        CanAssess = canAssess;
        CanRemove = canRemove;
        Notes = notes?.Trim() ?? string.Empty;
        BrowserLabel = browserLabel?.Trim();
        BrowserOrder = browserOrder;
    }

    /// <summary>Stable capability identifier.</summary>
    public string Id { get; }
    /// <summary>User-facing format label.</summary>
    public string Label { get; }
    /// <summary>Package that owns format-specific inspection and mutation semantics.</summary>
    public string OwnerPackage { get; }
    /// <summary>Exact extensions and structural format identities owned by this capability.</summary>
    public IReadOnlyList<OfficeProvenanceWorkflowFormat> Formats { get; }
    /// <summary>Recognized filename extensions.</summary>
    public IReadOnlyList<string> Extensions => _extensions;
    /// <summary>Extensions qualified for byte-oriented hosts.</summary>
    public IReadOnlyList<string> MemoryOnlyExtensions => _memoryOnlyExtensions;
    /// <summary>Whether structural inspection is supported.</summary>
    public bool CanInspect { get; }
    /// <summary>Whether combined assessment is supported.</summary>
    public bool CanAssess { get; }
    /// <summary>Whether selected carriers can be removed.</summary>
    public bool CanRemove { get; }
    /// <summary>Whether at least one format is qualified for the browser host.</summary>
    public bool BrowserAvailable => _browserAvailable;
    /// <summary>Compact label used by browser file-selection surfaces.</summary>
    public string? BrowserLabel { get; }
    /// <summary>Important capability boundary.</summary>
    public string Notes { get; }

    internal int BrowserOrder { get; }
    internal OfficeProvenanceWorkflowAdapter.ProvenanceOwner Owner { get; }

    /// <summary>Gets whether this capability supports an operation.</summary>
    public bool Supports(OfficeProvenanceWorkflowOperation operation) => operation switch {
        OfficeProvenanceWorkflowOperation.Inspect => CanInspect,
        OfficeProvenanceWorkflowOperation.Assess => CanAssess,
        OfficeProvenanceWorkflowOperation.Remove => CanRemove,
        _ => false
    };
}

/// <summary>Canonical OfficeIMO-owned format and operation catalog used by provenance workflow consumers.</summary>
public static partial class OfficeProvenanceWorkflowCatalog {
    /// <summary>Stable catalog identifier.</summary>
    public const string Id = "OfficeIMO.Provenance";
    /// <summary>Machine-readable catalog schema version.</summary>
    public const int SchemaVersion = 1;

    private static readonly IReadOnlyList<OfficeProvenanceWorkflowCapability> CapabilitiesValue =
        Array.AsReadOnly(new[] {
            Capability("word-openxml", "Word Open XML", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Word, PackageFormats(ModernExtensions(WordFormatCatalog.All), ".docx"), "Package signatures block mutation unless removal is explicitly authorized.", "DOCX", 2),
            Capability("excel-package", "Excel workbook package", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Excel, PackageFormats(ModernExtensions(ExcelFormatCatalog.All), ".xlsx"), "SpreadsheetML and XLSB package identity are validated before mutation.", "XLSX", 3),
            Capability("powerpoint-openxml", "PowerPoint Open XML", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.PowerPoint, PackageFormats(ModernExtensions(PowerPointFormatCatalog.All), ".pptx"), "Package signatures block mutation unless removal is explicitly authorized.", "PPTX", 4),
            Capability("visio-openxml", "Visio Open XML", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Visio, PackageFormats([".vsdm", ".vsdx", ".vssm", ".vssx", ".vstm", ".vstx"]), "Package signatures block mutation unless removal is explicitly authorized."),
            Capability("open-document", "OpenDocument", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.OpenDocument, PackageFormats([".odg", ".odp", ".ods", ".odt", ".otg", ".otp", ".ots", ".ott"]), "Encrypted OpenDocument packages cannot be rewritten."),
            Capability("epub", "EPUB", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Epub, PackageFormats([".epub"]), "Package structure is validated before inspection or mutation."),
            Capability("pdf", "PDF", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Pdf, Formats([".pdf"], OfficeProvenanceAssetFormat.Pdf, ".pdf"), "Removal is limited to provenance associations supported by the PDF owner.", "PDF", 1),
            Capability("html", "HTML", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Html, Formats([".htm", ".html"], OfficeProvenanceAssetFormat.Html), "External resources are not fetched during inspection."),
            Capability("markdown", "Markdown", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Markdown, StructuredTextFormats([".markdown", ".md"]), "Original BOM-aware UTF encoding is preserved by file mutation."),
            Capability("core-images", "Image provenance", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Core, [
                Format(".gif", OfficeProvenanceAssetFormat.Gif),
                Format(".jpeg", OfficeProvenanceAssetFormat.Jpeg, memoryOnly: true, browser: true),
                Format(".jpg", OfficeProvenanceAssetFormat.Jpeg, memoryOnly: true, browser: true),
                Format(".png", OfficeProvenanceAssetFormat.Png, memoryOnly: true, browser: true),
                Format(".svg", OfficeProvenanceAssetFormat.Svg),
                Format(".tif", OfficeProvenanceAssetFormat.Tiff),
                Format(".tiff", OfficeProvenanceAssetFormat.Tiff),
                Format(".webp", OfficeProvenanceAssetFormat.Webp, memoryOnly: true, browser: true)
            ], "File contents must match the registered image format; unrelated media containers are not inferred from signatures.", "JPEG / PNG / WebP", 0),
            Capability("core-text", "Structured text", OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Core, TextFormats([".adoc", ".asciidoc", ".bat", ".c", ".cjs", ".cmd", ".cpp", ".cs", ".css", ".go", ".h", ".hpp", ".ini", ".java", ".js", ".json", ".lua", ".mjs", ".ps1", ".py", ".rb", ".rs", ".sh", ".sql", ".tex", ".toml", ".ts", ".txt", ".vb", ".xml", ".yaml", ".yml"]), "Only standards-defined structured or wrapped text carriers are changed.")
        });

    private static readonly IReadOnlyDictionary<string, OfficeProvenanceWorkflowCapability> ByExtension =
        new ReadOnlyDictionary<string, OfficeProvenanceWorkflowCapability>(CapabilitiesValue
            .SelectMany(static capability => capability.Extensions.Select(extension => (extension, capability)))
            .ToDictionary(static item => item.extension, static item => item.capability, StringComparer.OrdinalIgnoreCase));
    private static readonly IReadOnlyDictionary<string, OfficeProvenanceWorkflowFormat> FormatsByExtension =
        new ReadOnlyDictionary<string, OfficeProvenanceWorkflowFormat>(CapabilitiesValue
            .SelectMany(static capability => capability.Formats)
            .ToDictionary(static item => item.Extension, StringComparer.OrdinalIgnoreCase));
    private static readonly IReadOnlyList<OfficeProvenanceWorkflowCapability> BrowserCapabilitiesValue =
        Array.AsReadOnly(CapabilitiesValue.Where(static item => item.BrowserAvailable).OrderBy(static item => item.BrowserOrder).ToArray());
    private static readonly IReadOnlyList<string> MemoryOnlyExtensionsValue = Array.AsReadOnly(CapabilitiesValue
        .SelectMany(static item => item.MemoryOnlyExtensions)
        .OrderBy(static item => item, StringComparer.Ordinal)
        .ToArray());
    private static readonly IReadOnlyList<string> BrowserExtensionsValue = Array.AsReadOnly(CapabilitiesValue
        .SelectMany(static item => item.Formats)
        .Where(static item => item.BrowserAvailable)
        .Select(static item => item.Extension)
        .OrderBy(static item => item, StringComparer.Ordinal)
        .ToArray());

    /// <summary>All OfficeIMO-owned cross-format provenance capabilities in stable order.</summary>
    public static IReadOnlyList<OfficeProvenanceWorkflowCapability> All => CapabilitiesValue;
    /// <summary>Capabilities qualified for the static browser host in presentation order.</summary>
    public static IReadOnlyList<OfficeProvenanceWorkflowCapability> BrowserCapabilities => BrowserCapabilitiesValue;
    /// <summary>All extensions qualified for byte-oriented provenance hosts.</summary>
    public static IReadOnlyList<string> MemoryOnlyExtensions => MemoryOnlyExtensionsValue;
    /// <summary>All extensions qualified for the static browser host.</summary>
    public static IReadOnlyList<string> BrowserExtensions => BrowserExtensionsValue;

    /// <summary>Finds the configured OfficeIMO owner by filename extension.</summary>
    public static OfficeProvenanceWorkflowCapability? FindByPath(string? path) {
        string extension = GetExtension(path);
        return ByExtension.TryGetValue(extension, out OfficeProvenanceWorkflowCapability? capability) ? capability : null;
    }

    /// <summary>Finds an exact registered format by filename extension.</summary>
    public static OfficeProvenanceWorkflowFormat? FindFormatByPath(string? path) {
        string extension = GetExtension(path);
        return FormatsByExtension.TryGetValue(extension, out OfficeProvenanceWorkflowFormat? format) ? format : null;
    }

    /// <summary>Finds a byte-oriented format only when that host has been qualified for it.</summary>
    public static OfficeProvenanceWorkflowFormat? FindMemoryOnlyFormatByPath(string? path) {
        OfficeProvenanceWorkflowFormat? format = FindFormatByPath(path);
        return format?.MemoryOnlyAvailable == true ? format : null;
    }

    private static OfficeProvenanceWorkflowCapability Capability(
        string id,
        string label,
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner owner,
        IEnumerable<OfficeProvenanceWorkflowFormat> formats,
        string notes,
        string? browserLabel = null,
        int browserOrder = int.MaxValue) =>
        new(id, label, owner, formats, canInspect: true, canAssess: true, canRemove: true, notes, browserLabel, browserOrder);

    private static IEnumerable<string> ModernExtensions(IReadOnlyList<OfficeFormatDescriptor> formats) =>
        formats.Where(static item => item.Generation == OfficeFormatGeneration.Modern).Select(static item => item.Extension);

    private static OfficeProvenanceWorkflowFormat[] PackageFormats(IEnumerable<string> extensions, string? browserExtension = null) =>
        Formats(extensions, OfficeProvenanceAssetFormat.ZipPackage, browserExtension);

    private static OfficeProvenanceWorkflowFormat[] TextFormats(IEnumerable<string> extensions) => extensions
        .Select(static extension => new OfficeProvenanceWorkflowFormat(
            extension,
            [OfficeProvenanceAssetFormat.StructuredText, OfficeProvenanceAssetFormat.UnstructuredText]))
        .ToArray();

    private static OfficeProvenanceWorkflowFormat[] StructuredTextFormats(IEnumerable<string> extensions) =>
        Formats(extensions, OfficeProvenanceAssetFormat.StructuredText);

    internal static string GetOwnerPackage(OfficeProvenanceWorkflowAdapter.ProvenanceOwner owner) => owner switch {
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Word => "OfficeIMO.Word",
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Excel => "OfficeIMO.Excel",
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner.PowerPoint => "OfficeIMO.PowerPoint",
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Visio => "OfficeIMO.Visio",
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner.OpenDocument => "OfficeIMO.OpenDocument",
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Epub => "OfficeIMO.Epub",
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Pdf => "OfficeIMO.Pdf",
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Html => "OfficeIMO.Html",
        OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Markdown => "OfficeIMO.Markdown",
        _ => "OfficeIMO.Core"
    };

    private static OfficeProvenanceWorkflowFormat[] Formats(
        IEnumerable<string> extensions,
        OfficeProvenanceAssetFormat assetFormat,
        string? browserExtension = null) => extensions
        .Select(extension => Format(
            extension,
            assetFormat,
            memoryOnly: string.Equals(extension, browserExtension, StringComparison.OrdinalIgnoreCase),
            browser: string.Equals(extension, browserExtension, StringComparison.OrdinalIgnoreCase)))
        .ToArray();

    private static OfficeProvenanceWorkflowFormat Format(
        string extension,
        OfficeProvenanceAssetFormat assetFormat,
        bool memoryOnly = false,
        bool browser = false) => new(extension, [assetFormat], memoryOnly, browser);

    private static string GetExtension(string? path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        string value = path.Trim();
        return value.StartsWith(".", StringComparison.Ordinal) && value.IndexOfAny(['/', '\\']) < 0
            ? value.ToLowerInvariant()
            : Path.GetExtension(value).ToLowerInvariant();
    }
}
