using Avalonia;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Editor;

/// <summary>A pending mark tied to the workspace revision from which its geometry was derived.</summary>
public sealed partial class PdfRedactionMarkViewModel : ObservableObject {
    internal PdfRedactionMarkViewModel(PdfRedactionArea area, Rect bounds, string description,
        PdfRedactionTextSelection textSelection = PdfRedactionTextSelection.LogicalBlocks) {
        Area = area;
        Bounds = bounds;
        Description = description;
        TextSelection = textSelection;
    }

    internal PdfRedactionArea Area { get; }
    internal Rect Bounds { get; }
    internal PdfRedactionTextSelection TextSelection { get; }
    public int PageNumber => Area.PageNumber;
    public string Description { get; }

    [ObservableProperty]
    private bool _isIncluded = true;

    [ObservableProperty]
    private string _reason = string.Empty;
}
