using OfficeIMO.Pdf;

namespace OfficeIMO.Web.Converter.Models;

public enum PdfToolInputMode {
    Single,
    Pair,
    Multiple
}

public enum PdfToolKind {
    Inspect,
    Merge,
    Split,
    Extract,
    Delete,
    Reorder,
    Rotate,
    Optimize,
    Protect,
    Unlock,
    Redact,
    Compare
}

public sealed record PdfToolDefinition(
    string Id,
    PdfToolKind Kind,
    string Group,
    string Label,
    string ShortLabel,
    string Description,
    PdfToolInputMode InputMode,
    string ActionLabel,
    bool RequiresPageSelection = false,
    bool RequiresPagesPerDocument = false,
    bool RequiresRotation = false,
    bool RequiresOptimizationProfile = false,
    bool RequiresUserPassword = false,
    bool RequiresOwnerPassword = false,
    bool RequiresRedactionText = false,
    bool RequiresDestructiveConfirmation = false);

public sealed record PdfToolRequest(
    PdfToolDefinition Tool,
    IReadOnlyList<SelectedDocument> Files,
    string PageSelection,
    int PagesPerDocument,
    int RotationDegrees,
    PdfOptimizationProfile OptimizationProfile,
    string UserPassword,
    string OwnerPassword,
    string RedactionText,
    bool DestructiveActionConfirmed);

public sealed record PdfToolMessage(string Title, string Message, string ToneClass);

public sealed record PdfFileMoveRequest(int Index, int Offset);

public sealed record PdfToolResult(
    BrowserConversionArtifact Artifact,
    BrowserConversionArtifact? Report,
    IReadOnlyList<PdfToolMessage> Messages,
    string Summary,
    int? PageCount,
    long SourceBytes,
    bool PreviewInBrowser);
