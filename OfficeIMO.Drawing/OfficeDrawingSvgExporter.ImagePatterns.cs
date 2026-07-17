using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendImagePattern(StringBuilder sb, OfficeDrawingImagePattern imagePattern, IOfficeRasterImageCodec? imageCodec, ref int elementId) {
        if (!OfficeSvgImageRenderer.TryCreateDataUri(imagePattern.ContentType, imagePattern.EncodedBytes, null, imageCodec, out string dataUri)) {
            return;
        }

        OfficeImagePatternLayout layout = imagePattern.Layout;
        OfficeImagePlacement area = layout.Area;
        OfficeImagePlacement tile = layout.Tile;
        double patternX = layout.RepeatX ? tile.X : area.X;
        double patternY = layout.RepeatY ? tile.Y : area.Y;
        double patternWidth = layout.RepeatX ? layout.HorizontalStep : area.Width;
        double patternHeight = layout.RepeatY ? layout.VerticalStep : area.Height;
        string patternId = "officeimo-image-pattern-" + (++elementId).ToString(CultureInfo.InvariantCulture);
        sb.Append("<defs><pattern")
            .AppendAttribute("id", patternId)
            .AppendAttribute("patternUnits", "userSpaceOnUse")
            .AppendAttribute("patternContentUnits", "userSpaceOnUse")
            .AppendNumberAttribute("x", patternX)
            .AppendNumberAttribute("y", patternY)
            .AppendNumberAttribute("width", patternWidth)
            .AppendNumberAttribute("height", patternHeight)
            .Append("><image")
            .AppendNumberAttribute("x", tile.X)
            .AppendNumberAttribute("y", tile.Y)
            .AppendNumberAttribute("width", tile.Width)
            .AppendNumberAttribute("height", tile.Height)
            .AppendAttribute("href", dataUri)
            .AppendAttribute("preserveAspectRatio", "none");
        if (imagePattern.Opacity < 1D) {
            sb.AppendNumberAttribute("opacity", imagePattern.Opacity);
        }

        sb.Append("/></pattern></defs><rect")
            .AppendNumberAttribute("x", area.X)
            .AppendNumberAttribute("y", area.Y)
            .AppendNumberAttribute("width", area.Width)
            .AppendNumberAttribute("height", area.Height)
            .AppendAttribute("fill", "url(#" + patternId + ")")
            .Append("/>");
    }
}
