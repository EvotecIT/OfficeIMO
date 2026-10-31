using OfficeIMO.Excel;

namespace PSWriteOffice.Models.Excel;

/// <summary>A worksheet view snapshot with its workbook path and sheet name.</summary>
public sealed class ExcelSheetViewInfo {
    internal ExcelSheetViewInfo(string sheetName, ExcelWorksheetViewInfo view, string? path) {
        SheetName = sheetName;
        Path = path;
        HasPane = view.HasPane;
        PaneState = view.PaneState;
        HorizontalSplit = view.HorizontalSplit;
        VerticalSplit = view.VerticalSplit;
        FrozenRowCount = view.FrozenRowCount;
        FrozenColumnCount = view.FrozenColumnCount;
        TopLeftCell = view.TopLeftCell;
        ActivePane = view.ActivePane;
        ShowGridlines = view.ShowGridlines;
        RightToLeft = view.RightToLeft;
        View = view.View;
        ZoomScale = view.ZoomScale;
        ZoomScaleNormal = view.ZoomScaleNormal;
    }
    /// <summary>Gets the worksheet name.</summary>
    public string SheetName { get; }
    /// <summary>Gets the worksheet name for sheet-based pipeline consumers.</summary>
    public string Sheet => SheetName;
    /// <summary>Gets the inspected file path, or null for a live workbook.</summary>
    public string? Path { get; }
    /// <summary>Gets whether the worksheet contains a pane.</summary>
    public bool HasPane { get; }
    /// <summary>Gets the native pane state, such as frozen or split.</summary>
    public string? PaneState { get; }
    /// <summary>Gets the native horizontal split.</summary>
    public double? HorizontalSplit { get; }
    /// <summary>Gets the native vertical split.</summary>
    public double? VerticalSplit { get; }
    /// <summary>Gets the number of frozen rows.</summary>
    public int FrozenRowCount { get; }
    /// <summary>Gets the number of frozen columns.</summary>
    public int FrozenColumnCount { get; }
    /// <summary>Gets the top-left scrollable cell.</summary>
    public string? TopLeftCell { get; }
    /// <summary>Gets the active pane.</summary>
    public string? ActivePane { get; }
    /// <summary>Gets whether gridlines are visible.</summary>
    public bool ShowGridlines { get; }
    /// <summary>Gets whether layout is right-to-left.</summary>
    public bool RightToLeft { get; }
    /// <summary>Gets the native view mode.</summary>
    public string? View { get; }
    /// <summary>Gets the current zoom percentage.</summary>
    public uint? ZoomScale { get; }
    /// <summary>Gets the normal-view zoom percentage.</summary>
    public uint? ZoomScaleNormal { get; }
}
