using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>
        ///     Places this text box, picture, table or chart into a placeholder of the slide's layout, the way
        ///     PowerPoint does when content is inserted into a placeholder, and keeps the placeholder binding.
        ///     A text box inherits the placeholder's position, size and text formatting from the layout, so later
        ///     layout edits still apply; a picture or chart takes the placeholder's position and size (pictures
        ///     are cropped to fill it), and a table spans its width with its own row heights. An empty slide instance of the placeholder is replaced.
        /// </summary>
        /// <param name="placeholderType">Layout placeholder type, such as <see cref="PowerPointPlaceholderType.Picture" /> or <see cref="PowerPointPlaceholderType.Object" />.</param>
        /// <param name="index">Placeholder index when the layout has several placeholders of the type.</param>
        public PowerPointShape BindToPlaceholder(PowerPointPlaceholderType placeholderType, uint? index = null) {
            PowerPointSlide slide = OwnerSlide ?? throw new InvalidOperationException("Only shapes on a slide can be placed into a layout placeholder.");
            slide.BindToPlaceholder(this, placeholderType, index);
            return this;
        }
    }

    public partial class PowerPointSlide {
        private const int BindingNamespace = OfficeXmlNamespaces.PresentationML;

        // The non-visual containers of a picture and a graphic frame, in sequence order.
        private static readonly OfficeXmlChildOrder PictureNonVisualOrder = new(
            new[] { BindingNamespace, BindingNamespace, BindingNamespace }, new[] { "cNvPr", "cNvPicPr", "nvPr" });
        private static readonly OfficeXmlChildOrder FrameNonVisualOrder = new(
            new[] { BindingNamespace, BindingNamespace, BindingNamespace }, new[] { "cNvPr", "cNvGraphicFramePr", "nvPr" });

        internal void BindToPlaceholder(PowerPointShape shape, PowerPointPlaceholderType placeholderType, uint? index) {
            EnsureShapeOnSlide(shape);
            if (shape.ShapeContentType is not (PowerPointShapeContentType.TextBox or PowerPointShapeContentType.Picture
                or PowerPointShapeContentType.Table or PowerPointShapeContentType.Chart)) {
                throw new NotSupportedException("Only text boxes, pictures, tables and charts can be placed into placeholders.");
            }

            PowerPointPartRef layout = PartRef.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
            if (layout.IsNone) throw new InvalidOperationException("The slide has no layout.");
            string wanted = placeholderType.ToOpenXmlToken();
            OfficeXmlNode layoutData = layout.Root.Child(BindingNamespace, "cSld");
            OfficeXmlNode layoutElement = default, layoutPlaceholder = default;
            foreach (OfficeXmlNode element in layoutData.Child(BindingNamespace, "spTree").Children()) {
                OfficeXmlNode candidate = PlaceholderOf(element);
                // The schema's default placeholder type is obj.
                if (candidate.IsNone || (candidate.Attribute("type") ?? "obj") != wanted
                    || index != null && (candidate.UInt32Attribute("idx") ?? 0U) != index.Value) continue;
                layoutElement = element;
                layoutPlaceholder = candidate;
                break;
            }
            if (layoutElement.IsNone) {
                throw new InvalidOperationException("Layout '" + layoutData.Attribute("name") + "' has no " + placeholderType
                    + " placeholder" + (index == null ? string.Empty : " with index " + index.Value) + ".");
            }
            uint layoutIndex = layoutPlaceholder.UInt32Attribute("idx") ?? 0U;

            PowerPointLayoutBox? bounds = ElementBounds(layoutElement) ?? GetExportPlaceholderBounds(layout, layoutPlaceholder);
            PowerPointShape? instance = ShapeList.FirstOrDefault(candidate => !ReferenceEquals(candidate, shape)
                && candidate.XmlElement is var element && element.Parent.Is(BindingNamespace, "spTree")
                && PlaceholderOf(element) is { IsNone: false } placeholder
                && (placeholder.UInt32Attribute("idx") ?? 0U) == layoutIndex
                && (layoutIndex != 0U || (placeholder.Attribute("type") ?? "obj") == wanted));
            OfficeXmlNode overrideTransform = default;
            if (instance != null) {
                if (instance is not PowerPointTextBox textBox || !string.IsNullOrEmpty(textBox.Text)) {
                    throw new InvalidOperationException("The " + placeholderType + " placeholder already holds content on this slide.");
                }

                // A slide-level override of the placeholder's position wins over the layout.
                OfficeXmlNode instanceElement = instance.XmlElement;
                bounds = ElementBounds(instanceElement) ?? bounds;
                if (instanceElement.Is(BindingNamespace, "sp")) {
                    OfficeXmlNode transform = instanceElement.Child(BindingNamespace, "spPr").Child(OfficeXmlNamespaces.DrawingML, "xfrm");
                    if (!transform.IsNone) overrideTransform = transform.Clone();
                }
            }

            if (!bounds.HasValue || bounds.Value.Width.ToEmus() <= 0 || bounds.Value.Height.ToEmus() <= 0) {
                throw new InvalidOperationException("The " + placeholderType + " placeholder has no usable bounds.");
            }

            if (instance != null) RemoveShape(instance);

            OfficeXmlNode target = shape.XmlElement;
            switch (target.LocalName) {
                case "sp": {
                    OfficeXmlNode nonVisual = target.GetOrAddChild(BindingNamespace, "nvSpPr", PShapeView.Order);
                    SetBinding(nonVisual.GetOrAddChild(BindingNamespace, "nvPr", PShapeNonVisualView.Order), layoutPlaceholder);
                    OfficeXmlNode drawing = nonVisual.GetOrAddChild(BindingNamespace, "cNvSpPr", PShapeNonVisualView.Order);
                    drawing.RemoveAttribute("txBox");
                    drawing.GetOrAddChild(OfficeXmlNamespaces.DrawingML, "spLocks", ANonVisualDrawingShapePropsView.Order).SetAttribute("noGrp", "1");
                    OfficeXmlNode properties = target.GetOrAddChild(BindingNamespace, "spPr", PShapeView.Order);
                    // Without a local transform the text follows the layout placeholder's geometry.
                    properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "xfrm");
                    if (!overrideTransform.IsNone) properties.InsertOrdered(overrideTransform, AShapePropertiesView.Order);
                    shape.OwnerPartRef.MarkEdited();
                    shape.OnGeometryChanged();
                    return;
                }
                case "pic": {
                    OfficeXmlNode nonVisual = target.GetOrPrependChild(BindingNamespace, "nvPicPr");
                    SetBinding(nonVisual.GetOrAddChild(BindingNamespace, "nvPr", PictureNonVisualOrder), layoutPlaceholder);
                    nonVisual.GetOrAddChild(BindingNamespace, "cNvPicPr", PictureNonVisualOrder)
                        .GetOrPrependChild(OfficeXmlNamespaces.DrawingML, "picLocks").SetAttribute("noGrp", "1");
                    break;
                }
                case "graphicFrame": {
                    OfficeXmlNode nonVisual = target.GetOrPrependChild(BindingNamespace, "nvGraphicFramePr");
                    SetBinding(nonVisual.GetOrAddChild(BindingNamespace, "nvPr", FrameNonVisualOrder), layoutPlaceholder);
                    nonVisual.GetOrAddChild(BindingNamespace, "cNvGraphicFramePr", FrameNonVisualOrder)
                        .GetOrPrependChild(OfficeXmlNamespaces.DrawingML, "graphicFrameLocks").SetAttribute("noGrp", "1");
                    break;
                }
            }
            shape.OwnerPartRef.MarkEdited();

            if (shape is PowerPointTable table) {
                // Like PowerPoint, a table spans the placeholder's width and keeps its own row heights.
                double[] ratios = Enumerable.Range(0, table.ColumnCount)
                    .Select(column => Math.Max(1D, table.GetColumnWidth(column).ToEmus())).ToArray();
                table.Left = bounds.Value.Left;
                table.Top = bounds.Value.Top;
                table.Width = bounds.Value.Width;
                if (ratios.Length > 0) table.SetColumnWidthsByRatio(ratios);
                return;
            }

            shape.Bounds = bounds.Value;
            if (shape is PowerPointImage image && image.TryGetPixelSize(out int width, out int height)) {
                image.FitToBox(width, height, crop: true);
            }
        }

        // Replaces the application properties' placeholder with one of the layout placeholder's type and index.
        private static void SetBinding(OfficeXmlNode application, OfficeXmlNode layoutPlaceholder) {
            application.RemoveChildren(BindingNamespace, "ph");
            OfficeXmlNode binding = application.InsertBefore(application.CreateElement(BindingNamespace, "ph"), application.FirstChild);
            if (layoutPlaceholder.Attribute("type") is { } type) binding.SetAttribute("type", type);
            if (layoutPlaceholder.Attribute("idx") is { } bindingIndex) binding.SetAttribute("idx", bindingIndex);
        }
    }
}
