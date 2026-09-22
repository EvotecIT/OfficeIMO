using System;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static void TryAddForeignObject(
        XElement element,
        OfficeDrawing drawing,
        SvgPaintContext style,
        SvgElementReferenceRegistry references,
        OfficeTransform transform,
        double viewX,
        double viewY,
        int maximumElements,
        ref int visited,
        ref int unsupported) {
        if (HasUnsupportedForeignObjectEffect(element)) unsupported++;
        if (style.Opacity <= 0D) return;

        OfficeSvgForeignObjectRenderer? renderer = references.ForeignObjectRenderer;
        if (renderer == null
            || !TryViewportLength(element, "width", drawing.Width, out double width)
            || !TryViewportLength(element, "height", drawing.Height, out double height)
            || width <= 0D
            || height <= 0D) {
            unsupported++;
            return;
        }

        double x = ReadViewportCoordinate(element, "x", viewX, drawing.Width);
        double y = ReadViewportCoordinate(element, "y", viewY, drawing.Height);
        if (!references.TryGetForeignObject(element, width, height, out OfficeDrawing content,
                out int contentElements, out double nestedPixels)) {
            if (!references.HasForeignObjectContent(element)) return;
            if (!references.TryReserveForeignObject(element, width, height)) {
                unsupported++;
                return;
            }
            string html = string.Concat(element.Nodes().Select(node => node.ToString(SaveOptions.DisableFormatting)));
            if (!references.TryChargeSerializedForeignObject(html.Length)) {
                unsupported++;
                return;
            }
            OfficeDrawing? rendered;
            try {
                rendered = renderer(new OfficeSvgForeignObjectContext(html, width, height));
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception exception) when (exception is not OutOfMemoryException && exception is not StackOverflowException) {
                unsupported++;
                return;
            }
            if (rendered == null
                || Math.Abs(rendered.Width - width) > 0.0001D
                || Math.Abs(rendered.Height - height) > 0.0001D) {
                unsupported++;
                return;
            }
            content = rendered;
            if (!TryMeasureForeignObjectSurfaces(content, maximumElements, out contentElements, out nestedPixels)) {
                unsupported++;
                return;
            }
            references.CacheForeignObject(element, width, height, content, contentElements, nestedPixels);
        }
        if (contentElements > maximumElements - visited) {
            unsupported++;
            return;
        }
        if (!references.TryChargeForeignObjectPlacement(drawing.Width, drawing.Height, nestedPixels)) {
            unsupported++;
            return;
        }
        visited += contentElements;

        var layer = new OfficeDrawing(drawing.Width, drawing.Height);
        layer.AddClippedDrawingForRendering(
            content,
            x,
            y,
            OfficeClipPath.Rectangle(width, height),
            contentOffsetX: 0D,
            contentOffsetY: 0D);
        drawing.AddEffectDrawing(layer, transform, style.Opacity);
    }

    private static bool HasUnsupportedForeignObjectEffect(XElement element) {
        string? filter = ReadPresentationProperty(element, "filter");
        string? mask = ReadPresentationProperty(element, "mask");
        string? blend = ReadPresentationProperty(element, "mix-blend-mode");
        return IsActiveEffectReference(filter)
            || IsActiveEffectReference(mask)
            || (!string.IsNullOrWhiteSpace(blend)
                && !blend!.Trim().Equals("normal", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsActiveEffectReference(string? value) {
        return !string.IsNullOrWhiteSpace(value)
            && !value!.Trim().Equals("none", StringComparison.OrdinalIgnoreCase);
    }
}
