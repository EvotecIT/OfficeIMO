using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

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
            string shapeName = name ?? GenerateUniqueName(shapeType.ToString());
            A.ShapeTypeValues openXmlShapeType = shapeType.ToOpenXml();

            ShapeProperties shapeProperties = new(
                new A.Transform2D(new A.Offset { X = nativeLeft, Y = nativeTop }, new A.Extents { Cx = nativeWidth, Cy = nativeHeight }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = openXmlShapeType }
            );

            if (shapeType == OfficePresetShapeType.Line) {
                shapeProperties.Append(new A.NoFill());
            }

            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = AllocateShapeId(), Name = shapeName },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                shapeProperties
            );

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            PowerPointAutoShape autoShape = TrackShape(new PowerPointAutoShape(shape));
            return autoShape;
        }

        internal PowerPointConnectionShape AddConnectionShape(A.ShapeTypeValues shapeType, long left,
            long top, long width, long height) {
            ShapeProperties shapeProperties = new(
                new A.Transform2D(new A.Offset { X = left, Y = top },
                    new A.Extents { Cx = width, Cy = height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = shapeType });
            var connection = new ConnectionShape(
                new NonVisualConnectionShapeProperties(
                    new NonVisualDrawingProperties {
                        Id = AllocateShapeId(),
                        Name = GenerateUniqueName("Connector")
                    },
                    new NonVisualConnectorShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                shapeProperties);
            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.Append(connection);
            return TrackShape(new PowerPointConnectionShape(connection));
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
