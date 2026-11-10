using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

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
        internal void BindToPlaceholder(PowerPointShape shape, PowerPointPlaceholderType placeholderType, uint? index) {
            EnsureShapeOnSlide(shape);
            if (shape.ShapeContentType is not (PowerPointShapeContentType.TextBox or PowerPointShapeContentType.Picture
                or PowerPointShapeContentType.Table or PowerPointShapeContentType.Chart)) {
                throw new NotSupportedException("Only text boxes, pictures, tables and charts can be placed into placeholders.");
            }

            SlideLayoutPart layoutPart = _slidePart.SlideLayoutPart
                ?? throw new InvalidOperationException("The slide has no layout.");
            PlaceholderValues wanted = placeholderType.ToOpenXml();
            OpenXmlElement layoutElement = layoutPart.SlideLayout?.CommonSlideData?.ShapeTree?.ChildElements
                .FirstOrDefault(element => GetLayoutPlaceholderShape(element) is { } candidate
                    && (candidate.Type?.Value ?? PlaceholderValues.Object) == wanted
                    && (index == null || (candidate.Index?.Value ?? 0U) == index.Value))
                ?? throw new InvalidOperationException("Layout '" + SlideLayoutName + "' has no " + placeholderType
                    + " placeholder" + (index == null ? string.Empty : " with index " + index.Value) + ".");
            PlaceholderShape layoutPlaceholder = GetLayoutPlaceholderShape(layoutElement)!;
            uint layoutIndex = layoutPlaceholder.Index?.Value ?? 0U;

            PowerPointLayoutBox? bounds = GetLayoutElementBounds(layoutElement)
                ?? GetExportPlaceholderBounds(layoutPart, layoutPlaceholder);
            PowerPointShape? instance = ShapeList.FirstOrDefault(candidate => !ReferenceEquals(candidate, shape)
                && candidate.Element.Parent is ShapeTree
                && GetLayoutPlaceholderShape(candidate.Element) is { } placeholder
                && (placeholder.Index?.Value ?? 0U) == layoutIndex
                && (layoutIndex != 0U || (placeholder.Type?.Value ?? PlaceholderValues.Object) == wanted));
            A.Transform2D? overrideTransform = null;
            if (instance != null) {
                if (instance is not PowerPointTextBox textBox || !string.IsNullOrEmpty(textBox.Text)) {
                    throw new InvalidOperationException("The " + placeholderType + " placeholder already holds content on this slide.");
                }

                // A slide-level override of the placeholder's position wins over the layout.
                bounds = GetLayoutElementBounds(instance.Element) ?? bounds;
                overrideTransform = (instance.Element as Shape)?.ShapeProperties?.Transform2D?.CloneNode(true) as A.Transform2D;
                RemoveShape(instance);
            }

            if (!bounds.HasValue || bounds.Value.Width.ToEmus() <= 0 || bounds.Value.Height.ToEmus() <= 0) {
                throw new InvalidOperationException("The " + placeholderType + " placeholder has no usable bounds.");
            }

            var binding = new PlaceholderShape();
            if (layoutPlaceholder.Type != null) binding.Type = layoutPlaceholder.Type.Value;
            if (layoutPlaceholder.Index != null) binding.Index = layoutPlaceholder.Index.Value;
            switch (shape.Element) {
                case Shape text:
                    NonVisualShapeProperties textProperties = text.NonVisualShapeProperties ??= new NonVisualShapeProperties();
                    (textProperties.ApplicationNonVisualDrawingProperties ??= new ApplicationNonVisualDrawingProperties()).PlaceholderShape = binding;
                    NonVisualShapeDrawingProperties textDrawing = textProperties.NonVisualShapeDrawingProperties
                        ??= new NonVisualShapeDrawingProperties();
                    textDrawing.TextBox = null;
                    (textDrawing.ShapeLocks ??= new A.ShapeLocks()).NoGrouping = true;
                    ShapeProperties textShapeProperties = text.ShapeProperties ??= new ShapeProperties();
                    // Without a local transform the text follows the layout placeholder's geometry.
                    textShapeProperties.Transform2D = overrideTransform;
                    shape.OnGeometryChanged();
                    return;
                case Picture picture:
                    NonVisualPictureProperties pictureProperties = picture.NonVisualPictureProperties ??= new NonVisualPictureProperties();
                    (pictureProperties.ApplicationNonVisualDrawingProperties ??= new ApplicationNonVisualDrawingProperties()).PlaceholderShape = binding;
                    NonVisualPictureDrawingProperties pictureDrawing = pictureProperties.NonVisualPictureDrawingProperties
                        ??= new NonVisualPictureDrawingProperties();
                    (pictureDrawing.PictureLocks ??= new A.PictureLocks()).NoGrouping = true;
                    break;
                case GraphicFrame frame:
                    NonVisualGraphicFrameProperties frameProperties = frame.NonVisualGraphicFrameProperties ??= new NonVisualGraphicFrameProperties();
                    (frameProperties.ApplicationNonVisualDrawingProperties ??= new ApplicationNonVisualDrawingProperties()).PlaceholderShape = binding;
                    NonVisualGraphicFrameDrawingProperties frameDrawing = frameProperties.NonVisualGraphicFrameDrawingProperties
                        ??= new NonVisualGraphicFrameDrawingProperties();
                    (frameDrawing.GraphicFrameLocks ??= new A.GraphicFrameLocks()).NoGrouping = true;
                    break;
            }

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
    }
}
