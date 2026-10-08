using Avalonia;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Editor;

/// <summary>A pending mark tied to the workspace revision from which its geometry was derived.</summary>
public sealed partial class PdfRedactionMarkViewModel : ObservableObject {
    internal PdfRedactionMarkViewModel(PdfRedactionArea area, Rect bounds, string description,
        PdfRedactionTextSelection textSelection = PdfRedactionTextSelection.LogicalBlocks,
        IStudioLocalizer? localizer = null) {
        Area = area;
        Bounds = bounds;
        Description = description;
        TextSelection = textSelection;
        localizer ??= new StudioLocalizer(System.Globalization.CultureInfo.GetCultureInfo("en"));
        string page = localizer.FormatOrDefault("PdfPage.Label", "Page {0}", PageNumber);
        AccessibleName = $"{page}: {description}";
        IncludeAccessibleName = localizer.FormatOrDefault("Redaction.IncludePage", "Include page {0}", PageNumber);
    }

    internal PdfRedactionArea Area { get; }
    internal Rect Bounds { get; }
    internal PdfRedactionTextSelection TextSelection { get; }
    public int PageNumber => Area.PageNumber;
    public string Description { get; }

    /// <summary>Identifies the page and matched content to assistive technology.</summary>
    public string AccessibleName { get; }

    /// <summary>Identifies the page controlled by the inclusion checkbox.</summary>
    public string IncludeAccessibleName { get; }

    [ObservableProperty]
    private bool _isIncluded = true;

    [ObservableProperty]
    private string _reason = string.Empty;
}
