namespace OfficeIMO.Excel.Xlsb.Model;

/// <summary>Settings of the primary window's worksheet view from BrtBeginWsView (MS-XLSB 2.4.307).</summary>
internal sealed class XlsbSheetView {
    private readonly List<XlsbSelection> _selections = new();

    internal XlsbSheetView(ushort flags, uint view, int topRow, int leftColumn, byte gridColor,
        ushort zoom, ushort zoomNormal, ushort zoomPageBreak, ushort zoomPageLayout) {
        Flags = flags;
        View = view;
        TopRow = topRow;
        LeftColumn = leftColumn;
        GridColor = gridColor;
        Zoom = zoom;
        ZoomNormal = zoomNormal;
        ZoomPageBreak = zoomPageBreak;
        ZoomPageLayout = zoomPageLayout;
    }

    /// <summary>fWnProt through fWhitespaceHidden, in the record's bit order.</summary>
    internal ushort Flags { get; }
    internal uint View { get; }
    internal int TopRow { get; }
    internal int LeftColumn { get; }
    internal byte GridColor { get; }
    internal ushort Zoom { get; }
    internal ushort ZoomNormal { get; }
    internal ushort ZoomPageBreak { get; }
    internal ushort ZoomPageLayout { get; }
    internal IReadOnlyList<XlsbSelection> Selections => _selections;

    internal void AddSelection(XlsbSelection selection) => _selections.Add(selection);
}

/// <summary>One BrtSel record of a worksheet view.</summary>
internal sealed class XlsbSelection {
    internal XlsbSelection(uint pane, int activeRow, int activeColumn, uint activeCellId, IReadOnlyList<XlsbCellRange> ranges) {
        Pane = pane;
        ActiveRow = activeRow;
        ActiveColumn = activeColumn;
        ActiveCellId = activeCellId;
        Ranges = ranges;
    }

    internal uint Pane { get; }
    /// <summary>One-based active row.</summary>
    internal int ActiveRow { get; }
    /// <summary>One-based active column.</summary>
    internal int ActiveColumn { get; }
    internal uint ActiveCellId { get; }
    internal IReadOnlyList<XlsbCellRange> Ranges { get; }
}
