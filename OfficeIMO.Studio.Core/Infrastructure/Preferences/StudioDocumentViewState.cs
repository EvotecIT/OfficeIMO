using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Infrastructure.Preferences;

internal sealed record StudioPanePreference(bool Navigation, bool Inspector);

/// <summary>Bounded presentation state; no document contents, passwords, or source paths are stored here.</summary>
internal sealed record StudioDocumentViewState {
    public double NavigationWidth { get; set; } = 238D;
    public double InspectorWidth { get; set; } = 300D;
    public int PageNumber { get; set; } = 1;
    public double Zoom { get; set; } = 1D;
    public ViewerZoomMode ZoomMode { get; set; } = ViewerZoomMode.FitWidth;
    public ReaderLayoutMode ReaderLayout { get; set; } = ReaderLayoutMode.Continuous;
    public bool FocusReading { get; set; }
    public Dictionary<StudioDocumentMode, StudioPanePreference> Panes { get; set; } = new();

    internal StudioDocumentViewState Normalize() => this with {
        NavigationWidth = double.IsFinite(NavigationWidth) ? Math.Clamp(NavigationWidth, 200D, 320D) : 238D,
        InspectorWidth = double.IsFinite(InspectorWidth) ? Math.Clamp(InspectorWidth, 280D, 380D) : 300D,
        PageNumber = Math.Clamp(PageNumber, 1, 1_000_000),
        Zoom = double.IsFinite(Zoom) ? Math.Clamp(Zoom, 0.25D, 3D) : 1D,
        ZoomMode = Enum.IsDefined(ZoomMode) ? ZoomMode : ViewerZoomMode.FitWidth,
        ReaderLayout = Enum.IsDefined(ReaderLayout) ? ReaderLayout : ReaderLayoutMode.Continuous,
        Panes = (Panes ?? new()).Where(pair => Enum.IsDefined(pair.Key) && pair.Value is not null)
            .ToDictionary(pair => pair.Key, pair => pair.Value)
    };
}
