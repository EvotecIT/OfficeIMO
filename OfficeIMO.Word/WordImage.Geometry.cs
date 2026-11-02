using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeIMO.Word;

public partial class WordImage {
    /// <summary>Gets or sets the displayed image width. Null means no width is stored.</summary>
    public OfficeLength? Width {
        get => ReadDisplayedLength(horizontal: true);
        set => WriteDisplayedLength(value ?? throw new ArgumentNullException(nameof(value)), horizontal: true);
    }

    /// <summary>Gets or sets the displayed image height. Null means no height is stored.</summary>
    public OfficeLength? Height {
        get => ReadDisplayedLength(horizontal: false);
        set => WriteDisplayedLength(value ?? throw new ArgumentNullException(nameof(value)), horizontal: false);
    }

    private OfficeLength? ReadDisplayedLength(bool horizontal) {
        if (_vmlShape != null) {
            double? pixels = ReadVmlDimension(horizontal ? "width" : "height");
            return pixels.HasValue ? OfficeLength.Pixels(pixels.Value) : null;
        }
        var extent = _Image.Inline?.Extent ?? _Image.Anchor?.Extent;
        long? emus = horizontal ? extent?.Cx?.Value : extent?.Cy?.Value;
        return emus.HasValue ? OfficeLength.Emus(emus.Value) : null;
    }

    private void WriteDisplayedLength(OfficeLength value, bool horizontal) {
        if (_vmlShape != null) {
            WriteVmlDimension(horizontal ? "width" : "height", value.ToPixels());
            return;
        }
        var extent = _Image.Inline?.Extent ?? _Image.Anchor?.Extent;
        if (extent == null) return;
        var graphic = _Image.Inline?.Extent != null
            ? _Image.Inline.Graphic
            : _Image.Anchor?.Elements<A.Graphic>().FirstOrDefault();
        var pictureExtent = graphic?.GraphicData?.GetFirstChild<Pic.Picture>()?.ShapeProperties?.Transform2D?.Extents;
        long emus = value.ToEmus();
        if (horizontal) {
            extent.Cx = emus;
            if (pictureExtent != null) pictureExtent.Cx = emus;
        } else {
            extent.Cy = emus;
            if (pictureExtent != null) pictureExtent.Cy = emus;
        }
    }
}
