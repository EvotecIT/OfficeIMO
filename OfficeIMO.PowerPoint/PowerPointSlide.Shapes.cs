using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds an auto shape with the specified geometry.
        /// </summary>
        public PowerPointAutoShape AddShape(OfficePresetShapeType shapeType, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null, string? name = null) {
            long nativeLeft = (left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(914400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(914400L)).ToEmus();
            string preset = shapeType.ToOpenXmlToken();
            if (TryAddPackageShape(shapeType, preset, nativeLeft, nativeTop, nativeWidth, nativeHeight, name) is { } added) {
                return added;
            }

            const int presentation = OfficeXmlNamespaces.PresentationML;
            uint id = AllocateShapeId();
            string shapeName = name ?? GenerateUniqueName(shapeType.ToString());
            PowerPointPartRef part = PartRef;
            OfficeXmlNode tree = EnsureCommonSlideData(part).GetOrAddChild(presentation, "spTree", PCommonSlideDataView.Order);
            OfficeXmlNode shape = CreatePresetShapeMarkup(tree, preset, shapeName, id, nativeLeft, nativeTop, nativeWidth, nativeHeight);
            if (shapeType == OfficePresetShapeType.Line) {
                OfficeXmlNode properties = shape.Child(presentation, "spPr");
                properties.Append(properties.CreateElement(OfficeXmlNamespaces.DrawingML, "noFill"));
            }

            shape = tree.Append(shape);
            part.MarkEdited();
            return TrackShape((PowerPointAutoShape)WrapSlideShape(shape, part)!);
        }

        private static OfficeXmlNode CreatePresetShapeMarkup(OfficeXmlNode tree, string preset, string name, uint id,
            long left, long top, long width, long height, bool isTextBox = false, bool isTitle = false) {
            const int presentation = OfficeXmlNamespaces.PresentationML;
            const int drawing = OfficeXmlNamespaces.DrawingML;
            OfficeXmlNode shape = tree.CreateElement(presentation, "sp");
            OfficeXmlNode nonVisual = shape.Append(shape.CreateElement(presentation, "nvSpPr"));
            OfficeXmlNode drawingProperties = nonVisual.Append(nonVisual.CreateElement(presentation, "cNvPr"));
            drawingProperties.SetUInt32Attribute("id", id);
            drawingProperties.SetAttribute("name", name);
            OfficeXmlNode shapeDrawing = nonVisual.Append(nonVisual.CreateElement(presentation, "cNvSpPr"));
            if (isTextBox) shapeDrawing.SetBooleanAttribute("txBox", true);
            shapeDrawing.Append(shapeDrawing.CreateElement(drawing, "spLocks")).SetBooleanAttribute("noGrp", true);
            OfficeXmlNode application = nonVisual.Append(nonVisual.CreateElement(presentation, "nvPr"));
            if (isTitle) application.Append(application.CreateElement(presentation, "ph")).SetAttribute("type", "title");

            OfficeXmlNode properties = shape.Append(shape.CreateElement(presentation, "spPr"));
            OfficeXmlNode transform = properties.Append(properties.CreateElement(drawing, "xfrm"));
            OfficeXmlNode offset = transform.Append(transform.CreateElement(drawing, "off"));
            offset.SetInt64Attribute("x", left);
            offset.SetInt64Attribute("y", top);
            OfficeXmlNode extents = transform.Append(transform.CreateElement(drawing, "ext"));
            extents.SetInt64Attribute("cx", width);
            extents.SetInt64Attribute("cy", height);
            OfficeXmlNode geometry = properties.Append(properties.CreateElement(drawing, "prstGeom"));
            geometry.SetAttribute("prst", preset);
            geometry.Append(geometry.CreateElement(drawing, "avLst"));
            return shape;
        }

        /// <summary>
        ///     Adds a rectangle shape.
        /// </summary>
        public PowerPointAutoShape AddRectangle(OfficeLength left, OfficeLength top, OfficeLength width,
            OfficeLength height, string? name = null) {
            return AddShape(OfficePresetShapeType.Rectangle, left, top, width, height, name);
        }

        /// <summary>
        ///     Adds an ellipse shape.
        /// </summary>
        public PowerPointAutoShape AddEllipse(OfficeLength left, OfficeLength top, OfficeLength width,
            OfficeLength height, string? name = null) {
            return AddShape(OfficePresetShapeType.Ellipse, left, top, width, height, name);
        }

        /// <summary>
        ///     Adds a line shape from the specified start and end points.
        /// </summary>
        public PowerPointAutoShape AddLine(OfficeLength startX, OfficeLength startY, OfficeLength endX,
            OfficeLength endY, string? name = null) {
            long nativeStartX = startX.ToEmus();
            long nativeStartY = startY.ToEmus();
            long nativeEndX = endX.ToEmus();
            long nativeEndY = endY.ToEmus();
            long left = Math.Min(nativeStartX, nativeEndX);
            long top = Math.Min(nativeStartY, nativeEndY);
            long width = Math.Abs(nativeEndX - nativeStartX);
            long height = Math.Abs(nativeEndY - nativeStartY);

            if (width == 0) width = 1;
            if (height == 0) height = 1;

            return AddShape(OfficePresetShapeType.Line, OfficeLength.Emus(left), OfficeLength.Emus(top),
                OfficeLength.Emus(width), OfficeLength.Emus(height), name);
        }
    }
}
