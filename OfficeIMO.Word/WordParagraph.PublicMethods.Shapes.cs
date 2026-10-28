namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Adds a VML rectangle to the paragraph and returns the created shape.</summary>
        /// <param name="width">Rectangle width.</param>
        /// <param name="height">Rectangle height.</param>
        /// <param name="fillColor">Fill color; white when omitted.</param>
        public WordShape AddShape(OfficeLength width, OfficeLength height, OfficeColor? fillColor = null) {
            return new WordShape(_document, this, width, height, (fillColor ?? OfficeColor.White).ToRgbHex());
        }

        /// <summary>Adds a VML shape to the paragraph and returns the created shape.</summary>
        /// <param name="shapeType">Type of shape to create.</param>
        /// <param name="width">Width, or the line end X coordinate.</param>
        /// <param name="height">Height, or the line end Y coordinate.</param>
        /// <param name="fillColor">Fill color; white when omitted.</param>
        /// <param name="strokeColor">Outline color; black when omitted.</param>
        /// <param name="strokeWeight">Outline thickness; one point when omitted.</param>
        /// <param name="arcSize">Dimensionless corner roundness fraction.</param>
        public WordShape AddShape(WordShapeType shapeType, OfficeLength width, OfficeLength height,
            OfficeColor? fillColor = null, OfficeColor? strokeColor = null, OfficeLength? strokeWeight = null, double arcSize = 0.25) {
            WordShape.ValidateVmlColors(fillColor, strokeColor);
            WordShape shape;
            switch (shapeType) {
                case WordShapeType.Rectangle:
                    shape = AddShape(width, height, fillColor);
                    break;
                case WordShapeType.Ellipse:
                    shape = WordShape.AddEllipse(this, width, height, fillColor);
                    break;
                case WordShapeType.RoundedRectangle:
                    shape = WordShape.AddRoundedRectangle(this, width, height, fillColor, arcSize);
                    break;
                case WordShapeType.Line:
                    return WordShape.AddLine(this, OfficeLength.Zero, OfficeLength.Zero, width, height, strokeColor, strokeWeight);
                default:
                    throw new ArgumentOutOfRangeException(nameof(shapeType), shapeType, null);
            }
            shape.Stroked = true;
            shape.StrokeColor = strokeColor ?? OfficeColor.Black;
            shape.StrokeWeight = strokeWeight ?? OfficeLength.Points(1);
            return shape;
        }

        /// <summary>Adds an inline DrawingML shape and returns the created shape.</summary>
        public WordShape AddShapeDrawing(WordShapeType shapeType, OfficeLength width, OfficeLength height) {
            return WordShape.AddDrawingShape(this, shapeType, width, height);
        }

        /// <summary>Adds a DrawingML shape anchored at an absolute position on the page.</summary>
        public WordShape AddShapeDrawing(WordShapeType shapeType, OfficeLength width, OfficeLength height, OfficeLength left, OfficeLength top) {
            return WordShape.AddDrawingShapeAnchored(this, shapeType, width, height, left, top);
        }

        /// <summary>Adds a VML line and returns the created line.</summary>
        /// <param name="startX">Start X coordinate.</param>
        /// <param name="startY">Start Y coordinate.</param>
        /// <param name="endX">End X coordinate.</param>
        /// <param name="endY">End Y coordinate.</param>
        /// <param name="color">Stroke color; black when omitted.</param>
        /// <param name="strokeWeight">Stroke thickness; one point when omitted.</param>
        public WordLine AddLine(OfficeLength startX, OfficeLength startY, OfficeLength endX, OfficeLength endY,
            OfficeColor? color = null, OfficeLength? strokeWeight = null) {
            return new WordLine(_document, this, startX, startY, endX, endY,
                (color ?? OfficeColor.Black).ToRgbHex(), strokeWeight ?? OfficeLength.Points(1));
        }
    }
}
