namespace OfficeIMO.Pdf;

/// <summary>
/// Simple viewer preferences emitted in the generated PDF catalog.
/// </summary>
public sealed class PdfViewerPreferencesOptions {
    /// <summary>Requests that viewers display the document title from metadata instead of the file name.</summary>
    public bool? DisplayDocTitle { get; set; }
    /// <summary>Requests that viewers hide the toolbar when the document is opened.</summary>
    public bool? HideToolbar { get; set; }
    /// <summary>Requests that viewers hide the menu bar when the document is opened.</summary>
    public bool? HideMenubar { get; set; }
    /// <summary>Requests that viewers hide user-interface elements when the document is opened.</summary>
    public bool? HideWindowUI { get; set; }
    /// <summary>Requests that viewers resize the window to fit the first displayed page.</summary>
    public bool? FitWindow { get; set; }
    /// <summary>Requests that viewers center the document window on screen.</summary>
    public bool? CenterWindow { get; set; }

    internal bool HasAny =>
        DisplayDocTitle.HasValue ||
        HideToolbar.HasValue ||
        HideMenubar.HasValue ||
        HideWindowUI.HasValue ||
        FitWindow.HasValue ||
        CenterWindow.HasValue;

    internal PdfViewerPreferencesOptions Clone() {
        return new PdfViewerPreferencesOptions {
            DisplayDocTitle = DisplayDocTitle,
            HideToolbar = HideToolbar,
            HideMenubar = HideMenubar,
            HideWindowUI = HideWindowUI,
            FitWindow = FitWindow,
            CenterWindow = CenterWindow
        };
    }
}
