using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds a title textbox to the slide.
        /// </summary>
        public PowerPointTextBox AddTitle(string text, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null) {
            ArgumentNullException.ThrowIfNull(text);
            if (!left.HasValue && !top.HasValue && !width.HasValue && !height.HasValue)
                return AddAuthoringTitle(text);
            PowerPointLayoutBox defaults = left.HasValue && top.HasValue && width.HasValue && height.HasValue
                ? new PowerPointLayoutBox(left.Value, top.Value, width.Value, height.Value) : GetAuthoringTitleBounds(text);
            long nativeLeft = (left ?? defaults.Left).ToEmus();
            long nativeTop = (top ?? defaults.Top).ToEmus();
            long nativeWidth = (width ?? defaults.Width).ToEmus();
            long nativeHeight = (height ?? defaults.Height).ToEmus();
            PreflightTitleReflow(top ?? defaults.Top, height ?? defaults.Height);

            PowerPointTextBox textBox = BuildTextShape(text, "rect", GenerateUniqueName("Title"), nativeLeft,
                nativeTop, nativeWidth, nativeHeight, isTitle: true);
            ReflowAuthoringContent();
            return textBox;
        }

        /// <summary>
        ///     Adds a title textbox using a layout box.
        /// </summary>
        public PowerPointTextBox AddTitle(string text, PowerPointLayoutBox layout) {
            return AddTitle(text, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds a textbox with the specified text.
        /// </summary>
        public PowerPointTextBox AddTextBox(string text, OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null,
            OfficeLength? height = null) {
            long nativeLeft = (left ?? OfficeLength.Emus(838200L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(2174875L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(7772400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(3962400L)).ToEmus();
            ArgumentNullException.ThrowIfNull(text);
            return BuildTextShape(text, "rect", GenerateUniqueName("TextBox"), nativeLeft,
                nativeTop, nativeWidth, nativeHeight, isTextBox: true);
        }

        private PowerPointTextBox BuildTextShape(string text, string preset, string name, long left, long top,
            long width, long height, bool isTextBox = false, bool isTitle = false) {
            const int presentation = OfficeXmlNamespaces.PresentationML;
            const int drawing = OfficeXmlNamespaces.DrawingML;
            uint id = AllocateShapeId();
            PowerPointPartRef part = PartRef;
            OfficeXmlNode tree = EnsureCommonSlideData(part).GetOrAddChild(presentation, "spTree", PCommonSlideDataView.Order);
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
            OfficeXmlNode body = shape.Append(shape.CreateElement(presentation, "txBody"));
            body.Append(body.CreateElement(drawing, "bodyPr"));
            body.Append(body.CreateElement(drawing, "lstStyle"));
            OfficeXmlNode paragraph = body.Append(body.CreateElement(drawing, "p"));
            if (paragraph.Tree is { } package) {
                PowerPointTreeText.AppendInlineContent(package, paragraph.Node, text, OfficeXmlTree.NoNode);
            } else {
                PowerPointTextContent.InsertInlineContent(paragraph, text, default, default);
            }

            shape = tree.Append(shape);
            part.MarkEdited();
            return TrackShape((PowerPointTextBox)WrapSlideShape(shape, part)!);
        }

        /// <summary>
        ///     Adds a textbox using a layout box.
        /// </summary>
        public PowerPointTextBox AddTextBox(string text, PowerPointLayoutBox layout) {
            return AddTextBox(text, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds an empty textbox using a layout box.
        /// </summary>
        public PowerPointTextBox AddTextBox(PowerPointLayoutBox layout) {
            return AddTextBox(string.Empty, layout.Left, layout.Top, layout.Width, layout.Height);
        }

    }
}
