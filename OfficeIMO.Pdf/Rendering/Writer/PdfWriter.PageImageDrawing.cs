using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void AppendPageImageDraw(StringBuilder sb, PageImage img) {
        if (img.ClipPath != null) {
            new ContentStreamBuilder(sb)
                .SaveState();
            AppendClipPath(sb, img.ClipPath, img.ClipX, img.ClipY, img.ClipHeight);
        }

        bool hasAlternativeText = !img.SuppressAccessibilityWrapper && !string.IsNullOrWhiteSpace(img.AlternativeText);
        if (hasAlternativeText) {
            sb.Append("/Figure << /Alt ")
                .Append(PdfSyntaxEscaper.TextString(img.AlternativeText!));
            if (img.MarkedContentId.HasValue) {
                sb.Append(" /MCID ")
                    .Append(img.MarkedContentId.Value.ToString(CultureInfo.InvariantCulture));
            }

            sb.Append(" >> BDC\n");
        } else if (!img.SuppressAccessibilityWrapper && img.IsDecorativeArtifact) {
            sb.Append("/Artifact << /Type /Layout >> BDC\n");
        }

        OfficeTransform imageTransform = new OfficeImageProjection(
            new OfficeImagePlacement(img.X, img.Y, img.W, img.H),
            rotationDegrees: img.RotationAngle,
            rotationCenterX: img.RotationCenterX,
            rotationCenterY: img.RotationCenterY,
            flipHorizontal: img.HorizontalFlip,
            flipVertical: img.VerticalFlip)
            .CreateUnitSquareTransform();

        if (img.IsForeground) imageTransform = imageTransform.Then(img.ForegroundTransform);
        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (!string.IsNullOrEmpty(img.GraphicsStateName)) {
            content.GraphicsState(img.GraphicsStateName!);
        }

        content
            .TransformMatrix(imageTransform);
        if (img.SourceCrop?.HasCrop == true) {
            double clipWidth = 1D - img.SourceCrop.Left - img.SourceCrop.Right;
            double clipHeight = 1D - img.SourceCrop.Top - img.SourceCrop.Bottom;
            content.Rectangle(img.SourceCrop.Left, img.SourceCrop.Bottom, clipWidth, clipHeight)
                .ClipPath()
                .EndPath();
        }

        if (img.Orientation != OfficeImageOrientation.Normal) {
            content.TransformMatrix(GetImageOrientationTransform(img.Orientation));
        }
        content.XObject(img.Name)
            .RestoreState();

        if (hasAlternativeText || !img.SuppressAccessibilityWrapper && img.IsDecorativeArtifact) {
            sb.Append("EMC\n");
        }

        if (img.ClipPath != null) {
            new ContentStreamBuilder(sb)
                .RestoreState();
        }
    }

}
