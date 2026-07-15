namespace OfficeIMO.Pdf;

/// <summary>Controls viewer, print, export, and locking behavior for generated optional content.</summary>
public sealed class PdfLayerOptions {
    /// <summary>Whether the layer is visible in the default configuration.</summary>
    public bool InitiallyVisible { get; set; } = true;
    /// <summary>Whether compatible viewers should prevent users from toggling the layer.</summary>
    public bool Locked { get; set; }
    /// <summary>Whether the layer is intended to be visible on screen.</summary>
    public bool VisibleInViewer { get; set; } = true;
    /// <summary>Whether the layer is intended to be visible when printed.</summary>
    public bool VisibleWhenPrinting { get; set; } = true;
    /// <summary>Whether the layer is intended to be visible when exported.</summary>
    public bool VisibleWhenExporting { get; set; } = true;

    internal PdfLayerOptions Clone() => new PdfLayerOptions {
        InitiallyVisible = InitiallyVisible,
        Locked = Locked,
        VisibleInViewer = VisibleInViewer,
        VisibleWhenPrinting = VisibleWhenPrinting,
        VisibleWhenExporting = VisibleWhenExporting
    };
}
