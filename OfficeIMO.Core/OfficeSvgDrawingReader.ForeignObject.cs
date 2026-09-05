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
        string html = string.Concat(element.Nodes().Select(node => node.ToString(SaveOptions.DisableFormatting)));
        if (string.IsNullOrWhiteSpace(html)) return;

        OfficeDrawing? content;
        try {
            content = renderer(new OfficeSvgForeignObjectContext(html, width, height));
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) when (exception is not OutOfMemoryException && exception is not StackOverflowException) {
            unsupported++;
            return;
        }

        if (content == null
            || Math.Abs(content.Width - width) > 0.0001D
            || Math.Abs(content.Height - height) > 0.0001D) {
            unsupported++;
            return;
        }

        int contentElements = CountDrawingElements(content, maximumElements);
        if (contentElements > maximumElements - visited) {
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
}
