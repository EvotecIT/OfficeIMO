using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>Shared drawing settings for exporting a recovered publication page as SVG.</summary>
public sealed class PublisherSvgOptions {
    /// <summary>Scale of the physical output surface. Drawing coordinates remain in points.</summary>
    public double Scale { get; set; } = 1;
    /// <summary>Unit of the root SVG dimensions. The default is physical points.</summary>
    public OfficeSvgSizeUnit SizeUnit { get; set; } = OfficeSvgSizeUnit.Point;
    /// <summary>Optional caller codec for image types outside the shared managed rendering contract.</summary>
    public IOfficeRasterImageCodec? ImageCodec { get; set; }
    /// <summary>Optional prefix for resource IDs when composing several inline SVG documents.</summary>
    public string? ResourceIdPrefix { get; set; }
}
