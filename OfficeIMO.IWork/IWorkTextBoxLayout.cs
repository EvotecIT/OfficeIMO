namespace OfficeIMO.IWork;

/// <summary>Vertical alignment of text inside a native text frame.</summary>
public enum IWorkTextVerticalAlignment {
    /// <summary>Text starts at the top of the frame.</summary>
    Top,
    /// <summary>Text is centered vertically.</summary>
    Middle,
    /// <summary>Text ends at the bottom of the frame.</summary>
    Bottom
}

/// <summary>Selected text-frame formatting, including inherited padding and alignment.</summary>
public sealed class IWorkTextBoxLayout {
    internal IWorkTextBoxLayout(bool? shrinkToFit, IWorkTextVerticalAlignment? verticalAlignment,
        double? left, double? top, double? right, double? bottom, bool singleColumn = false) {
        ShrinkToFit = shrinkToFit;
        VerticalAlignment = verticalAlignment;
        LeftInsetPoints = left;
        TopInsetPoints = top;
        RightInsetPoints = right;
        BottomInsetPoints = bottom;
        SupportsFixedFramePagination = singleColumn && shrinkToFit.HasValue && verticalAlignment.HasValue
            && left.HasValue && top.HasValue && right.HasValue && bottom.HasValue;
    }

    internal bool SupportsFixedFramePagination { get; }

    /// <summary>Gets whether text scales down to fit its fixed frame; null means unspecified.</summary>
    public bool? ShrinkToFit { get; }
    /// <summary>Gets the selected vertical alignment.</summary>
    public IWorkTextVerticalAlignment? VerticalAlignment { get; }
    /// <summary>Gets left text padding in points.</summary>
    public double? LeftInsetPoints { get; }
    /// <summary>Gets top text padding in points.</summary>
    public double? TopInsetPoints { get; }
    /// <summary>Gets right text padding in points.</summary>
    public double? RightInsetPoints { get; }
    /// <summary>Gets bottom text padding in points.</summary>
    public double? BottomInsetPoints { get; }
}
