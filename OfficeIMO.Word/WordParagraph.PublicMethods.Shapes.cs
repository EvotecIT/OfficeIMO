using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Reflection;
using System.Linq;
using System.Xml.Linq;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word {
    /// <summary>
    /// Contains public methods for editing paragraphs.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Adds a legacy VML image to the paragraph.
        /// </summary>
        public WordParagraph AddImageVml(string filePathImage, double? width = null, double? height = null) {
            var run = this.VerifyRun();
            MainDocumentPart mainPart = _document._wordprocessingDocument.MainDocumentPart!;

            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var fs = File.OpenRead(filePathImage)) {
                imagePart.FeedData(fs);
            }
            var relId = mainPart.GetIdOfPart(imagePart);

            string style = "mso-wrap-style:square";
            if (width.HasValue) style = $"width:{width}pt;" + style;
            if (height.HasValue) style = $"height:{height}pt;" + style;

            var shape = new V.Shape() {
                Id = "Image" + Guid.NewGuid().ToString("N"),
                Style = style,
                Type = "#_x0000_t75"
            };
            var imageData = new V.ImageData() {
                RelationshipId = relId,
                Title = Path.GetFileName(filePathImage)
            };
            shape.Append(imageData);
            Picture pict = new Picture();
            pict.Append(shape);
            run.Append(pict);
            return this;
        }

        /// <summary>
        /// Add a rectangle shape to the paragraph.
        /// </summary>
        /// <param name="widthPt">Width in points.</param>
        /// <param name="heightPt">Height in points.</param>
        /// <param name="fillColor">Fill color in hex format.</param>
        public WordShape AddShape(double widthPt, double heightPt, string fillColor = "#FFFFFF") {
            WordShape wordShape = new WordShape(this._document, this, widthPt, heightPt, fillColor);
            return wordShape;
        }

        /// <summary>
        /// Add a rectangle shape to the paragraph using <see cref="OfficeIMO.Drawing.OfficeColor"/>.
        /// </summary>
        public WordShape AddShape(double widthPt, double heightPt, OfficeIMO.Drawing.OfficeColor fillColor) {
            return AddShape(widthPt, heightPt, fillColor.ToRgbHex());
        }

        /// <summary>
        /// Adds a basic shape to the paragraph.
        /// </summary>
        /// <param name="shapeType">Type of shape to create.</param>
        /// <param name="widthPt">Width in points or line end X.</param>
        /// <param name="heightPt">Height in points or line end Y.</param>
        /// <param name="fillColor">Fill color in hex format.</param>
        /// <param name="strokeColor">Stroke color in hex format.</param>
        /// <param name="strokeWeightPt">Stroke weight in points.</param>
        /// <param name="arcSize">Corner roundness fraction for rounded rectangles.</param>
        public WordShape AddShape(ShapeType shapeType, double widthPt, double heightPt,
            string fillColor = "#FFFFFF", string strokeColor = "#000000", double strokeWeightPt = 1, double arcSize = 0.25) {
            WordShape shape;
            switch (shapeType) {
                case ShapeType.Rectangle:
                    shape = AddShape(widthPt, heightPt, fillColor);
                    break;
                case ShapeType.Ellipse:
                    shape = WordShape.AddEllipse(this, widthPt, heightPt, fillColor);
                    break;
                case ShapeType.RoundedRectangle:
                    shape = WordShape.AddRoundedRectangle(this, widthPt, heightPt, fillColor, arcSize);
                    break;
                case ShapeType.Line:
                    shape = WordShape.AddLine(this, 0, 0, widthPt, heightPt, strokeColor, strokeWeightPt);
                    return shape;
                default:
                    throw new ArgumentOutOfRangeException(nameof(shapeType), shapeType, null);
            }

            shape.Stroked = true;
            shape.StrokeColorHex = strokeColor;
            shape.StrokeWeight = strokeWeightPt;
            return shape;
        }

        /// <summary>
        /// Adds a basic shape to the paragraph using <see cref="OfficeIMO.Drawing.OfficeColor"/> values.
        /// </summary>
        public WordShape AddShape(ShapeType shapeType, double widthPt, double heightPt,
            OfficeIMO.Drawing.OfficeColor fillColor, OfficeIMO.Drawing.OfficeColor strokeColor, double strokeWeightPt = 1, double arcSize = 0.25) {
            return AddShape(shapeType, widthPt, heightPt, fillColor.ToRgbHex(), strokeColor.ToRgbHex(), strokeWeightPt, arcSize);
        }

        /// <summary>
        /// Adds a DrawingML shape to the paragraph.
        /// </summary>
        /// <param name="shapeType">Type of shape to create.</param>
        /// <param name="widthPt">Width in points.</param>
        /// <param name="heightPt">Height in points.</param>
        public WordShape AddShapeDrawing(ShapeType shapeType, double widthPt, double heightPt) {
            return WordShape.AddDrawingShape(this, shapeType, widthPt, heightPt);
        }

        /// <summary>
        /// Adds a DrawingML shape anchored at an absolute position on the page.
        /// </summary>
        /// <param name="shapeType">Type of shape to create.</param>
        /// <param name="widthPt">Width in points.</param>
        /// <param name="heightPt">Height in points.</param>
        /// <param name="leftPt">Left offset from page in points.</param>
        /// <param name="topPt">Top offset from page in points.</param>
        public WordShape AddShapeDrawing(ShapeType shapeType, double widthPt, double heightPt, double leftPt, double topPt) {
            return WordShape.AddDrawingShapeAnchored(this, shapeType, widthPt, heightPt, leftPt, topPt);
        }

        /// <summary>
        /// Add a line shape to the paragraph.
        /// </summary>
        /// <param name="startXPt">Start X position in points.</param>
        /// <param name="startYPt">Start Y position in points.</param>
        /// <param name="endXPt">End X position in points.</param>
        /// <param name="endYPt">End Y position in points.</param>
        /// <param name="color">Stroke color in hex format.</param>
        /// <param name="strokeWeightPt">Stroke weight in points.</param>
        public WordLine AddLine(double startXPt, double startYPt, double endXPt, double endYPt, string color = "#000000", double strokeWeightPt = 1) {
            var v = color;
            if (!string.IsNullOrEmpty(v) && !v.StartsWith("#", StringComparison.Ordinal)) v = "#" + v;
            WordLine wordLine = new WordLine(this._document, this, startXPt, startYPt, endXPt, endYPt, v, strokeWeightPt);
            return wordLine;
        }

        /// <summary>
        /// Add a line shape to the paragraph using <see cref="OfficeIMO.Drawing.OfficeColor"/>.
        /// </summary>
        public WordLine AddLine(double startXPt, double startYPt, double endXPt, double endYPt, OfficeIMO.Drawing.OfficeColor color, double strokeWeightPt = 1) {
            return AddLine(startXPt, startYPt, endXPt, endYPt, color.ToRgbHex(), strokeWeightPt);
        }
    }
}