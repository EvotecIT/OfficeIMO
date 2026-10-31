using System.Globalization;
using DocumentFormat.OpenXml;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordImage {
    /// <summary>Preserves the original picture's placement while its cropped frame changes.</summary>
    private (long X, long Y)? GetPhysicalCropAnchorPosition(int? left, int? top, int? right, int? bottom,
        long width, long height) {
        DW.Anchor? anchor = _Image.Anchor;
        if (anchor == null) return null;
        double oldWidth = ((double?)Width?.ToEmus())!.Value, oldHeight = ((double?)Height?.ToEmus())!.Value;
        double sourceWidth = oldWidth / VisibleCropFraction(CropLeftPercentage, CropRightPercentage);
        double sourceHeight = oldHeight / VisibleCropFraction(CropTopPercentage, CropBottomPercentage);
        double dx = ((left ?? 0D) - (CropLeftPercentage ?? 0D) - (right ?? 0D) + (CropRightPercentage ?? 0D))
            / CropPercentageUnits * sourceWidth / 2D;
        double dy = ((top ?? 0D) - (CropTopPercentage ?? 0D) - (bottom ?? 0D) + (CropBottomPercentage ?? 0D))
            / CropPercentageUnits * sourceHeight / 2D;
        if (HorizontalFlip == true) dx = -dx;
        if (VerticalFlip == true) dy = -dy;
        double angle = (Rotation ?? 0) * Math.PI / 180D;
        double deltaX = Math.Cos(angle) * dx - Math.Sin(angle) * dy + (oldWidth - width) / 2D;
        double deltaY = Math.Sin(angle) * dx + Math.Cos(angle) * dy + (oldHeight - height) / 2D;
        // Native Word crops the transformed source about its center, then updates the frame origin.
        if (anchor.SimplePos?.Value == true) {
            long x = anchor.SimplePosition?.X?.Value ?? throw InvalidPhysicalCropAnchor();
            long y = anchor.SimplePosition?.Y?.Value ?? throw InvalidPhysicalCropAnchor();
            return (AddPhysicalCropOffset(x, deltaX), AddPhysicalCropOffset(y, deltaY));
        }
        double horizontal = GetPhysicalCropAxisOffset(anchor.HorizontalPosition, true, oldWidth);
        double vertical = GetPhysicalCropAxisOffset(anchor.VerticalPosition, false, oldHeight);
        return (AddPhysicalCropOffset(horizontal, deltaX), AddPhysicalCropOffset(vertical, deltaY));
    }

    private double GetPhysicalCropAxisOffset(OpenXmlCompositeElement? position, bool horizontal, double frameExtent) {
        if (position == null) throw InvalidPhysicalCropAnchor();
        if (position.ChildElements.Any(child => child is not DW.PositionOffset &&
                child is not DW.HorizontalAlignment && child is not DW.VerticalAlignment))
            throw InvalidPhysicalCropAnchor();
        DW.PositionOffset? offset = position.GetFirstChild<DW.PositionOffset>();
        OpenXmlLeafTextElement? alignment = horizontal
            ? position.GetFirstChild<DW.HorizontalAlignment>()
            : position.GetFirstChild<DW.VerticalAlignment>();
        if (offset != null) {
            if (alignment != null || !ReadPositionOffset(offset).HasValue) throw InvalidPhysicalCropAnchor();
            return ReadPositionOffset(offset)!.Value.ToEmus();
        }
        if (alignment == null) throw InvalidPhysicalCropAnchor();
        string reference = position.GetAttribute("relativeFrom", string.Empty).Value ?? throw InvalidPhysicalCropAnchor();
        double area = GetPhysicalCropAlignmentArea(reference, horizontal);
        return alignment.Text switch {
            "left" when horizontal => 0D,
            "top" when !horizontal => 0D,
            "center" => (area - frameExtent) / 2D,
            "right" when horizontal => area - frameExtent,
            "bottom" when !horizontal => area - frameExtent,
            _ => throw InvalidPhysicalCropAnchor()
        };
    }

    private double GetPhysicalCropAlignmentArea(string reference, bool horizontal) {
        W.Paragraph? paragraph = _Image.Ancestors<W.Paragraph>().FirstOrDefault();
        WordSection? section = paragraph == null ? null : new WordParagraph(_document, paragraph, newRun: false).GetBodySection();
        if (section == null && _document.Sections.Count == 1) section = _document.Sections[0];
        if (section == null) throw InvalidPhysicalCropAnchor();
        W.PageSize? page = section._sectionProperties.GetFirstChild<W.PageSize>();
        double pageExtent = horizontal ? page?.Width?.Value ?? 0D : page?.Height?.Value ?? 0D;
        if (!IsPositiveFiniteCropValue(pageExtent)) throw InvalidPhysicalCropAnchor();
        if (reference == "page") return pageExtent * 635D;
        // Alignment to a margin is resolvable from section geometry; line, character,
        // multi-column and inside/outside alignment require a paginated layout context.
        W.PageMargin? margins = section._sectionProperties.GetFirstChild<W.PageMargin>();
        W.Settings? settings = _document._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings;
        if (margins?.Gutter?.Value > 0 || settings?.GetFirstChild<W.MirrorMargins>() is { } mirrored && mirrored.Val?.Value != false)
            throw InvalidPhysicalCropAnchor();
        double first = horizontal ? checked((uint)(section.Margins.Left.GetValueOrDefault().ToTwips())) : checked((int?)(section.Margins.Top?.ToTwips())) ?? 0D;
        double last = horizontal ? checked((uint)(section.Margins.Right.GetValueOrDefault().ToTwips())) : checked((int?)(section.Margins.Bottom?.ToTwips())) ?? 0D;
        if (reference == "column" && horizontal) {
            W.Columns? columns = section._sectionProperties.GetFirstChild<W.Columns>();
            if (columns?.ColumnCount?.Value > 1 || columns?.ChildElements.Count > 1) throw InvalidPhysicalCropAnchor();
            reference = "margin";
        }
        double extent = reference switch {
            "margin" => pageExtent - first - last,
            "leftMargin" when horizontal => first,
            "rightMargin" when horizontal => last,
            "topMargin" when !horizontal => first,
            "bottomMargin" when !horizontal => last,
            _ => throw InvalidPhysicalCropAnchor()
        };
        if (!IsPositiveFiniteCropValue(extent)) throw InvalidPhysicalCropAnchor();
        return extent * 635D;
    }

    private static long AddPhysicalCropOffset(double original, double delta) {
        double value = Math.Round(original + delta, MidpointRounding.AwayFromZero);
        if (double.IsNaN(value) || double.IsInfinity(value) || value < int.MinValue || value > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(delta), "The cropped image position must fit the DrawingML anchor range.");
        return checked((long)value);
    }

    private void ApplyPhysicalCropAnchorPosition((long X, long Y)? position) {
        if (!position.HasValue) return;
        DW.Anchor anchor = _Image.Anchor!;
        if (anchor.SimplePos?.Value == true) {
            anchor.SimplePosition!.X = position.Value.X;
            anchor.SimplePosition.Y = position.Value.Y;
        } else {
            anchor.HorizontalPosition!.RemoveAllChildren();
            anchor.HorizontalPosition.AppendChild(new DW.PositionOffset(position.Value.X.ToString(CultureInfo.InvariantCulture)));
            anchor.VerticalPosition!.RemoveAllChildren();
            anchor.VerticalPosition.AppendChild(new DW.PositionOffset(position.Value.Y.ToString(CultureInfo.InvariantCulture)));
        }
    }

    private static InvalidOperationException InvalidPhysicalCropAnchor() => new InvalidOperationException(
        "Physical cropping requires explicit image offsets or page/margin alignment that can be resolved from the section. Use raw percentage crop properties when position requires pagination.");
}
