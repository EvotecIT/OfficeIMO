namespace OfficeIMO.IWork;

/// <summary>Page dimensions and margins recovered from a Pages document.</summary>
public sealed class IWorkPageLayout {
    internal IWorkPageLayout(double widthPoints, double heightPoints, double leftMarginPoints,
        double rightMarginPoints, double topMarginPoints, double bottomMarginPoints,
        double headerMarginPoints, double footerMarginPoints, bool isLandscape) {
        WidthPoints = widthPoints;
        HeightPoints = heightPoints;
        LeftMarginPoints = leftMarginPoints;
        RightMarginPoints = rightMarginPoints;
        TopMarginPoints = topMarginPoints;
        BottomMarginPoints = bottomMarginPoints;
        HeaderMarginPoints = headerMarginPoints;
        FooterMarginPoints = footerMarginPoints;
        IsLandscape = isLandscape;
    }

    /// <summary>Gets page width in points.</summary>
    public double WidthPoints { get; }
    /// <summary>Gets page height in points.</summary>
    public double HeightPoints { get; }
    /// <summary>Gets left margin in points.</summary>
    public double LeftMarginPoints { get; }
    /// <summary>Gets right margin in points.</summary>
    public double RightMarginPoints { get; }
    /// <summary>Gets top margin in points.</summary>
    public double TopMarginPoints { get; }
    /// <summary>Gets bottom margin in points.</summary>
    public double BottomMarginPoints { get; }
    /// <summary>Gets header distance in points.</summary>
    public double HeaderMarginPoints { get; }
    /// <summary>Gets footer distance in points.</summary>
    public double FooterMarginPoints { get; }
    /// <summary>Gets whether the source declares landscape orientation.</summary>
    public bool IsLandscape { get; }
}

/// <summary>Position and size recovered from an iWork drawable, in points.</summary>
public sealed class IWorkGeometry {
    internal IWorkGeometry(double leftPoints, double topPoints, double widthPoints,
        double heightPoints, double rotationDegrees) {
        LeftPoints = leftPoints;
        TopPoints = topPoints;
        WidthPoints = widthPoints;
        HeightPoints = heightPoints;
        RotationDegrees = rotationDegrees;
    }

    /// <summary>Gets the horizontal position in points.</summary>
    public double LeftPoints { get; }
    /// <summary>Gets the vertical position in points.</summary>
    public double TopPoints { get; }
    /// <summary>Gets the width in points.</summary>
    public double WidthPoints { get; }
    /// <summary>Gets the height in points.</summary>
    public double HeightPoints { get; }
    /// <summary>Gets clockwise rotation in degrees.</summary>
    public double RotationDegrees { get; }
}

/// <summary>Two-dimensional canvas size recovered from an iWork document.</summary>
public sealed class IWorkCanvasSize {
    internal IWorkCanvasSize(double widthPoints, double heightPoints) {
        WidthPoints = widthPoints;
        HeightPoints = heightPoints;
    }
    /// <summary>Gets width in points.</summary>
    public double WidthPoints { get; }
    /// <summary>Gets height in points.</summary>
    public double HeightPoints { get; }
}

/// <summary>An embedded image recovered from an iWork data reference.</summary>
public sealed class IWorkImageAsset {
    private readonly byte[] _bytes;

    internal IWorkImageAsset(string fileName, string packagePath, string mediaType,
        byte[] bytes, int? pixelWidth, int? pixelHeight, IWorkGeometry? geometry,
        bool hasMask, string? hyperlink, string? accessibilityDescription) {
        FileName = fileName;
        PackagePath = packagePath;
        MediaType = mediaType;
        _bytes = bytes.ToArray();
        PixelWidth = pixelWidth;
        PixelHeight = pixelHeight;
        Geometry = geometry;
        HasMask = hasMask;
        Hyperlink = hyperlink;
        AccessibilityDescription = accessibilityDescription;
    }

    /// <summary>Gets the source-preferred file name.</summary>
    public string FileName { get; }
    /// <summary>Gets the normalized package entry path.</summary>
    public string PackagePath { get; }
    /// <summary>Gets the detected media type.</summary>
    public string MediaType { get; }
    /// <summary>Gets the embedded byte count.</summary>
    public long Length => _bytes.LongLength;
    /// <summary>Gets validated raster width, when available.</summary>
    public int? PixelWidth { get; }
    /// <summary>Gets validated raster height, when available.</summary>
    public int? PixelHeight { get; }
    /// <summary>Gets source drawable geometry.</summary>
    public IWorkGeometry? Geometry { get; }
    /// <summary>Gets whether the source image has a mask/crop object.</summary>
    public bool HasMask { get; }
    /// <summary>Gets a drawable hyperlink target.</summary>
    public string? Hyperlink { get; }
    /// <summary>Gets source accessibility description text.</summary>
    public string? AccessibilityDescription { get; }
    /// <summary>Returns a defensive copy of the embedded bytes.</summary>
    public byte[] GetBytes() => _bytes.ToArray();
}

/// <summary>A positioned rich-text box recovered from an iWork canvas.</summary>
public sealed class IWorkTextBox {
    internal IWorkTextBox(IWorkTextContent content, IWorkGeometry? geometry,
        string? hyperlink, string? accessibilityDescription) {
        Content = content;
        Geometry = geometry;
        Hyperlink = hyperlink;
        AccessibilityDescription = accessibilityDescription;
    }

    /// <summary>Gets the rich text content.</summary>
    public IWorkTextContent Content { get; }
    /// <summary>Gets source drawable geometry.</summary>
    public IWorkGeometry? Geometry { get; }
    /// <summary>Gets a drawable hyperlink target.</summary>
    public string? Hyperlink { get; }
    /// <summary>Gets source accessibility description text.</summary>
    public string? AccessibilityDescription { get; }
}
