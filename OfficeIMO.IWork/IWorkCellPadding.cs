namespace OfficeIMO.IWork;

/// <summary>Selected native table-cell padding in points. A declared empty padding message sets every side to zero.</summary>
public sealed class IWorkCellPadding {
    internal IWorkCellPadding(double left, double top, double right, double bottom) {
        LeftPoints = left;
        TopPoints = top;
        RightPoints = right;
        BottomPoints = bottom;
    }

    /// <summary>Gets the nonnegative left inset in points.</summary>
    public double LeftPoints { get; }
    /// <summary>Gets the nonnegative top inset in points.</summary>
    public double TopPoints { get; }
    /// <summary>Gets the nonnegative right inset in points.</summary>
    public double RightPoints { get; }
    /// <summary>Gets the nonnegative bottom inset in points.</summary>
    public double BottomPoints { get; }
}

/// <summary>The selected native vertical position of content within a table cell.</summary>
public enum IWorkCellVerticalAlignment {
    /// <summary>Content sits at the top of the cell.</summary>
    Top,
    /// <summary>Content is centered vertically.</summary>
    Middle,
    /// <summary>Content sits at the bottom of the cell.</summary>
    Bottom
}
