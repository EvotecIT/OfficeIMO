using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

internal static class PdfToolCatalog {
    internal static IReadOnlyList<PdfToolDefinition> All { get; } = [
        new("inspect", PdfToolKind.Inspect, "Understand", "Inspect PDF", "Inspect", "Read pages, security, forms, signatures, active content, and rewrite readiness without changing the document.", PdfToolInputMode.Single, "Inspect PDF"),
        new("compare", PdfToolKind.Compare, "Understand", "Compare PDFs", "Compare", "Render and compare two PDFs locally, then review expected, actual, and highlighted page differences.", PdfToolInputMode.Pair, "Compare PDFs"),
        new("merge", PdfToolKind.Merge, "Organize", "Merge PDFs", "Merge", "Combine two or more PDFs in the selected order through the first-party merge engine.", PdfToolInputMode.Multiple, "Merge PDFs"),
        new("split", PdfToolKind.Split, "Organize", "Split PDF", "Split", "Create consecutive PDF parts and download them together as a ZIP archive.", PdfToolInputMode.Single, "Split PDF", RequiresPagesPerDocument: true),
        new("extract", PdfToolKind.Extract, "Organize", "Extract pages", "Extract", "Create a new PDF from one-based page selections such as 1-3,5,last.", PdfToolInputMode.Single, "Extract pages", RequiresPageSelection: true),
        new("delete", PdfToolKind.Delete, "Organize", "Delete pages", "Delete", "Create a new PDF without the selected pages; the source file is never changed.", PdfToolInputMode.Single, "Delete pages", RequiresPageSelection: true, RequiresDestructiveConfirmation: true),
        new("reorder", PdfToolKind.Reorder, "Organize", "Reorder pages", "Reorder", "Create a new PDF with every page copied in the supplied order.", PdfToolInputMode.Single, "Reorder pages", RequiresPageSelection: true),
        new("rotate", PdfToolKind.Rotate, "Organize", "Rotate pages", "Rotate", "Rotate selected pages by 90, 180, or 270 degrees while preserving the original file.", PdfToolInputMode.Single, "Rotate pages", RequiresPageSelection: true, RequiresRotation: true),
        new("optimize", PdfToolKind.Optimize, "Publish", "Optimize PDF", "Optimize", "Apply deterministic lossless compression, deduplication, or Fast Web View without rasterizing pages.", PdfToolInputMode.Single, "Optimize PDF", RequiresOptimizationProfile: true),
        new("protect", PdfToolKind.Protect, "Secure", "Protect PDF", "Protect", "Encrypt a PDF with AES-256 Standard security and return preservation evidence.", PdfToolInputMode.Single, "Protect PDF", RequiresUserPassword: true, RequiresOwnerPassword: true),
        new("unlock", PdfToolKind.Unlock, "Secure", "Unlock PDF", "Unlock", "Remove Standard password security with the owner password and return a separate unprotected artifact.", PdfToolInputMode.Single, "Unlock PDF", RequiresOwnerPassword: true),
        new("redact", PdfToolKind.Redact, "Secure", "Redact text", "Redact", "Permanently remove literal text matches and verify that the marker is absent from rewritten content and streams.", PdfToolInputMode.Single, "Redact and verify", RequiresRedactionText: true, RequiresDestructiveConfirmation: true)
    ];

    internal static PdfToolDefinition Default => All[0];

    internal static PdfToolDefinition Find(string? id) =>
        All.FirstOrDefault(tool => string.Equals(tool.Id, id, StringComparison.OrdinalIgnoreCase)) ?? Default;
}
