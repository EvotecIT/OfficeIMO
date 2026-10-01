namespace OfficeIMO.IWork;

/// <summary>A supported selected table-cell fill, including an explicit no-fill override.</summary>
public sealed class IWorkCellFill {
    internal IWorkCellFill(IWorkColor? color) => Color = color;

    /// <summary>Gets the opaque source color. Null represents an explicit no-fill declaration.</summary>
    public IWorkColor? Color { get; }

    /// <summary>Gets whether the selected style explicitly clears an inherited fill.</summary>
    public bool IsNone => Color == null;
}
