using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using A = DocumentFormat.OpenXml.Drawing;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

#nullable enable annotations

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents simple shapes inside a paragraph.
    /// </summary>
    public class WordShape : WordElement {
        private const int EmusPerPoint = 12700; // 1 pt = 12700 EMUs
        private static int _docPrIdSeed = 1;
        private static UInt32Value NextDocPrId() {
            int id = Interlocked.Increment(ref _docPrIdSeed);
            return (UInt32Value)(uint)id;
        }
        /// <summary>Parent document.</summary>
        internal WordDocument _document = null!;
        /// <summary>Parent paragraph.</summary>
        internal WordParagraph _wordParagraph = null!;
        /// <summary>Run that hosts the shape.</summary>
        internal Run _run = null!;
        /// <summary>The rectangle element if present.</summary>
        internal V.Rectangle? _rectangle;
        /// <summary>The rounded rectangle element if present.</summary>
        internal V.RoundRectangle? _roundRectangle;
        /// <summary>The ellipse element if present.</summary>
        internal V.Oval? _ellipse;
        /// <summary>The line element if present.</summary>
        internal V.Line? _line;
        /// <summary>The polygon element if present.</summary>
        internal V.PolyLine? _polygon;
        /// <summary>The generic shape element if present.</summary>
        internal V.Shape? _shape;
        /// <summary>DrawingML shape element if present.</summary>
        internal Drawing? _drawing;
        internal Wps.WordprocessingShape? _wpsShape;
        
        private static (A.ShapeTypeValues preset, A.AdjustValueList adjustList) MapPresetGeometry(ShapeType shapeType) {
            var adjustList = new A.AdjustValueList();
            switch (shapeType) {
                case ShapeType.Line:
                    return (A.ShapeTypeValues.Line, adjustList);
                case ShapeType.Ellipse:
                    return (A.ShapeTypeValues.Ellipse, adjustList);
                case ShapeType.Rectangle:
                    return (A.ShapeTypeValues.Rectangle, adjustList);
                case ShapeType.RoundedRectangle:
                    adjustList.Append(new A.ShapeGuide() { Name = "adj", Formula = "val 16667" });
                    return (A.ShapeTypeValues.RoundRectangle, adjustList);
                case ShapeType.Triangle:
                    return (A.ShapeTypeValues.Triangle, adjustList);
                case ShapeType.Diamond:
                    return (A.ShapeTypeValues.Diamond, adjustList);
                case ShapeType.Pentagon:
                    return (A.ShapeTypeValues.Pentagon, adjustList);
                case ShapeType.Hexagon:
                    return (A.ShapeTypeValues.Hexagon, adjustList);
                case ShapeType.RightArrow:
                    return (A.ShapeTypeValues.RightArrow, adjustList);
                case ShapeType.LeftArrow:
                    return (A.ShapeTypeValues.LeftArrow, adjustList);
                case ShapeType.UpArrow:
                    return (A.ShapeTypeValues.UpArrow, adjustList);
                case ShapeType.DownArrow:
                    return (A.ShapeTypeValues.DownArrow, adjustList);
                case ShapeType.Star5:
                    return (A.ShapeTypeValues.Star5, adjustList);
                case ShapeType.Heart:
                    return (A.ShapeTypeValues.Heart, adjustList);
                case ShapeType.Cloud:
                    return (A.ShapeTypeValues.Cloud, adjustList);
                case ShapeType.Donut:
                    return (A.ShapeTypeValues.Donut, adjustList);
                case ShapeType.Can:
                    return (A.ShapeTypeValues.Can, adjustList);
                case ShapeType.Cube:
                    return (A.ShapeTypeValues.Cube, adjustList);
                default:
                    throw new ArgumentOutOfRangeException(nameof(shapeType), shapeType, null);
            }
        }

        private static long ToEmuChecked(double points, string paramName)
        {
            // reject NaN/Infinity and absurd negatives early
            if (double.IsNaN(points) || double.IsInfinity(points))
                throw new ArgumentOutOfRangeException(paramName, "Value must be a finite number.");
            // Convert using checked bounds to avoid long overflow
            double emu = points * EmusPerPoint;
            if (emu > long.MaxValue || emu < long.MinValue)
                throw new ArgumentOutOfRangeException(paramName, "Value is too large for EMU conversion.");
            return (long)Math.Round(emu);
        }

        private static void ValidateDimensions(double widthPt, double heightPt)
        {
            if (widthPt <= 0) throw new ArgumentOutOfRangeException(nameof(widthPt), "Width must be positive.");
            if (heightPt <= 0) throw new ArgumentOutOfRangeException(nameof(heightPt), "Height must be positive.");
        }

        private static void ValidatePosition(double leftPt, double topPt)
        {
            if (leftPt < 0) throw new ArgumentOutOfRangeException(nameof(leftPt), "Left offset cannot be negative.");
            if (topPt < 0) throw new ArgumentOutOfRangeException(nameof(topPt), "Top offset cannot be negative.");
        }

        /// <summary>
        /// Initializes a new rectangle shape and appends it to the paragraph.
        /// </summary>
        internal WordShape(WordDocument document, WordParagraph paragraph, double widthPt, double heightPt, string fillColor = "#FFFFFF") {
            _document = document;
            _wordParagraph = paragraph;

            // Ensure VML color has leading '#'
            var vmlFill = fillColor;
            if (!string.IsNullOrEmpty(vmlFill) && !vmlFill.StartsWith("#", StringComparison.Ordinal)) vmlFill = "#" + vmlFill;

            _rectangle = new V.Rectangle() {
                Id = "Rectangle" + Guid.NewGuid().ToString("N"),
                Style = $"width:{widthPt}pt;height:{heightPt}pt;mso-wrap-style:square",
                FillColor = vmlFill,
                Stroked = false
            };

            Picture pict = new Picture();
            pict.Append(_rectangle);

            _run = paragraph.VerifyRun();
            _run.Append(pict);
        }

        /// <summary>
        /// Initializes a <see cref="WordShape"/> from existing run content.
        /// </summary>
        internal WordShape(WordDocument document, Paragraph paragraph, Run run, Drawing? drawing = null) {
            _document = document;
            _wordParagraph = new WordParagraph(document, paragraph, run);
            _run = run;
            _rectangle = run.Descendants<V.Rectangle>().FirstOrDefault();
            _roundRectangle = run.Descendants<V.RoundRectangle>().FirstOrDefault();
            _ellipse = run.Descendants<V.Oval>().FirstOrDefault();
            _line = run.Descendants<V.Line>().FirstOrDefault();
            _polygon = run.Descendants<V.PolyLine>().FirstOrDefault();
            _shape = run.Descendants<V.Shape>().FirstOrDefault(s => !s.Descendants<V.ImageData>().Any() && !s.Descendants<V.TextBox>().Any());
            _drawing = drawing;
            if (drawing != null) {
                _wpsShape = drawing.Descendants<Wps.WordprocessingShape>().FirstOrDefault();
            }
        }

        /// <summary>
        /// Adds an ellipse shape to the given paragraph.
        /// </summary>
        public static WordShape AddEllipse(WordParagraph paragraph, double widthPt, double heightPt, string fillColor = "#FFFFFF") {
            var vmlFill = fillColor;
            if (!string.IsNullOrEmpty(vmlFill) && !vmlFill.StartsWith("#", StringComparison.Ordinal)) vmlFill = "#" + vmlFill;
            var ellipse = new V.Oval() {
                Id = "Ellipse" + Guid.NewGuid().ToString("N"),
                Style = $"width:{widthPt}pt;height:{heightPt}pt;mso-wrap-style:square",
                FillColor = vmlFill,
                Stroked = false
            };

            Picture pict = new Picture();
            pict.Append(ellipse);

            var run = paragraph.VerifyRun();
            run.Append(pict);

            return new WordShape(paragraph._document!, paragraph._paragraph!, run);
        }

        /// <summary>
        /// Adds an ellipse shape using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public static WordShape AddEllipse(WordParagraph paragraph, double widthPt, double heightPt, SixLabors.ImageSharp.Color fillColor) {
            return AddEllipse(paragraph, widthPt, heightPt, fillColor.ToHexColor());
        }

        /// <summary>
        /// Adds a rounded rectangle shape to the given paragraph.
        /// </summary>
        public static WordShape AddRoundedRectangle(WordParagraph paragraph, double widthPt, double heightPt,
            string fillColor = "#FFFFFF", double arcSize = 0.25) {
            var arc = (int)Math.Round(arcSize * 65536d);
            var vmlFill = fillColor;
            if (!string.IsNullOrEmpty(vmlFill) && !vmlFill.StartsWith("#", StringComparison.Ordinal)) vmlFill = "#" + vmlFill;
            var roundRect = new V.RoundRectangle() {
                Id = "RoundedRect" + Guid.NewGuid().ToString("N"),
                Style = $"width:{widthPt}pt;height:{heightPt}pt;mso-wrap-style:square",
                FillColor = vmlFill,
                Stroked = false,
                ArcSize = $"{arc}f"
            };

            Picture pict = new Picture();
            pict.Append(roundRect);

            var run = paragraph.VerifyRun();
            run.Append(pict);

            return new WordShape(paragraph._document!, paragraph._paragraph!, run);
        }

        /// <summary>
        /// Adds a line shape to the given paragraph.
        /// </summary>
        public static WordShape AddLine(WordParagraph paragraph, double startXPt, double startYPt, double endXPt, double endYPt, string color = "#000000", double strokeWeightPt = 1) {
            var vmlStroke = color;
            if (!string.IsNullOrEmpty(vmlStroke) && !vmlStroke.StartsWith("#", StringComparison.Ordinal)) vmlStroke = "#" + vmlStroke;
            var line = new V.Line() {
                Id = "Line" + Guid.NewGuid().ToString("N"),
                Style = "mso-wrap-style:square",
                From = $"{startXPt}pt,{startYPt}pt",
                To = $"{endXPt}pt,{endYPt}pt",
                StrokeColor = vmlStroke,
                StrokeWeight = $"{strokeWeightPt}pt"
            };

            Picture pict = new Picture();
            pict.Append(line);

            var run = paragraph.VerifyRun();
            run.Append(pict);

            return new WordShape(paragraph._document!, paragraph._paragraph!, run);
        }

        /// <summary>
        /// Adds a line shape using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public static WordShape AddLine(WordParagraph paragraph, double startXPt, double startYPt, double endXPt, double endYPt, SixLabors.ImageSharp.Color color, double strokeWeightPt = 1) {
            return AddLine(paragraph, startXPt, startYPt, endXPt, endYPt, color.ToHexColor(), strokeWeightPt);
        }

        /// <summary>
        /// Adds a polygon shape to the given paragraph.
        /// </summary>
        public static WordShape AddPolygon(WordParagraph paragraph, string points, string fillColor = "#FFFFFF", string strokeColor = "#000000") {
            var vmlFill = fillColor; if (!string.IsNullOrEmpty(vmlFill) && !vmlFill.StartsWith("#", StringComparison.Ordinal)) vmlFill = "#" + vmlFill;
            var vmlStroke = strokeColor; if (!string.IsNullOrEmpty(vmlStroke) && !vmlStroke.StartsWith("#", StringComparison.Ordinal)) vmlStroke = "#" + vmlStroke;
            var poly = new V.PolyLine() {
                Id = "Polygon" + Guid.NewGuid().ToString("N"),
                Style = "mso-wrap-style:square",
                Points = points,
                FillColor = vmlFill,
                StrokeColor = vmlStroke
            };

            Picture pict = new Picture();
            pict.Append(poly);

            var run = paragraph.VerifyRun();
            run.Append(pict);

            return new WordShape(paragraph._document!, paragraph._paragraph!, run);
        }

        /// <summary>
        /// Adds a polygon shape using <see cref="SixLabors.ImageSharp.Color"/> values.
        /// </summary>
        public static WordShape AddPolygon(WordParagraph paragraph, string points, SixLabors.ImageSharp.Color fillColor, SixLabors.ImageSharp.Color strokeColor) {
            return AddPolygon(paragraph, points, fillColor.ToHexColor(), strokeColor.ToHexColor());
        }

        /// <summary>
        /// Adds a DrawingML shape to the given paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph to append the shape to.</param>
        /// <param name="shapeType">Type of shape to create.</param>
        /// <param name="widthPt">Width in points.</param>
        /// <param name="heightPt">Height in points.</param>
        public static WordShape AddDrawingShape(WordParagraph paragraph, ShapeType shapeType, double widthPt, double heightPt) {
            ValidateDimensions(widthPt, heightPt);
            long cx = ToEmuChecked(widthPt, nameof(widthPt));
            long cy = ToEmuChecked(heightPt, nameof(heightPt));

            var run = paragraph.VerifyRun();

            var inline = new DW.Inline() {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            };

            inline.Append(new DW.Extent() { Cx = cx, Cy = cy });
            inline.Append(new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });
            inline.Append(new DW.DocProperties() { Id = NextDocPrId(), Name = "Shape" });
            inline.Append(new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }));

            var graphic = new A.Graphic();
            var graphicData = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
            var wsp = new Wps.WordprocessingShape();
            wsp.Append(new Wps.NonVisualDrawingShapeProperties(new A.ShapeLocks() { NoChangeArrowheads = true }));

            // Use Wps.ShapeProperties (wps:spPr) per schema; do not emit a:spPr directly under wps:wsp
            var shapeProps = new Wps.ShapeProperties();
            shapeProps.Append(new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = cx, Cy = cy }));

            var (preset, adjustList) = MapPresetGeometry(shapeType);

            shapeProps.Append(new A.PresetGeometry(adjustList) { Preset = preset });
            wsp.Append(shapeProps);

            var textBodyProps = new Wps.TextBodyProperties() {
                Rotation = 0,
                UseParagraphSpacing = false,
                VerticalOverflow = A.TextVerticalOverflowValues.Overflow,
                HorizontalOverflow = A.TextHorizontalOverflowValues.Overflow,
                Vertical = A.TextVerticalValues.Horizontal,
                Wrap = A.TextWrappingValues.Square,
                LeftInset = 91440,
                TopInset = 45720,
                RightInset = 91440,
                BottomInset = 45720,
                ColumnCount = 1,
                ColumnSpacing = 0,
                RightToLeftColumns = false,
                FromWordArt = false,
                Anchor = A.TextAnchoringTypeValues.Center,
                AnchorCenter = false,
                ForceAntiAlias = false,
                CompatibleLineSpacing = true
            };
            // Basic text behavior (no auto-fit). Word tolerates missing preset text wrap.
            textBodyProps.Append(new A.NoAutoFit());

            wsp.Append(textBodyProps);

            graphicData.Append(wsp);
            graphic.Append(graphicData);
            inline.Append(graphic);

            var drawing = new Drawing(inline);
            run.Append(drawing);

            return new WordShape(paragraph._document!, paragraph._paragraph!, run, drawing);
        }

        /// <summary>
        /// Adds a DrawingML shape anchored at an absolute position on the page.
        /// </summary>
        /// <param name="paragraph">Paragraph to host the drawing anchor.</param>
        /// <param name="shapeType">Type of shape.</param>
        /// <param name="widthPt">Width in points.</param>
        /// <param name="heightPt">Height in points.</param>
        /// <param name="leftPt">Left offset in points from the page.</param>
        /// <param name="topPt">Top offset in points from the page.</param>
        public static WordShape AddDrawingShapeAnchored(WordParagraph paragraph, ShapeType shapeType, double widthPt, double heightPt, double leftPt, double topPt) {
            ValidateDimensions(widthPt, heightPt);
            ValidatePosition(leftPt, topPt);
            long cx = ToEmuChecked(widthPt, nameof(widthPt));
            long cy = ToEmuChecked(heightPt, nameof(heightPt));
            long offX = ToEmuChecked(leftPt, nameof(leftPt));
            long offY = ToEmuChecked(topPt, nameof(topPt));

            var run = paragraph.VerifyRun();

            var anchor = new DW.Anchor() {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 0U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };

            anchor.Append(new DW.SimplePosition() { X = 0L, Y = 0L });
            var hpos = new DW.HorizontalPosition() { RelativeFrom = DW.HorizontalRelativePositionValues.Page };
            hpos.Append(new DW.PositionOffset(offX.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            anchor.Append(hpos);
            var vpos = new DW.VerticalPosition() { RelativeFrom = DW.VerticalRelativePositionValues.Page };
            vpos.Append(new DW.PositionOffset(offY.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            anchor.Append(vpos);
            anchor.Append(new DW.Extent() { Cx = cx, Cy = cy });
            anchor.Append(new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });
            anchor.Append(new DW.WrapSquare() { WrapText = DW.WrapTextValues.BothSides });
            anchor.Append(new DW.DocProperties() { Id = NextDocPrId(), Name = "Shape" });
            anchor.Append(new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }));

            var graphic = new A.Graphic();
            var graphicData = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };
            var wsp = new Wps.WordprocessingShape();
            wsp.Append(new Wps.NonVisualDrawingShapeProperties(new A.ShapeLocks() { NoChangeArrowheads = true }));

            var shapeProps = new Wps.ShapeProperties();
            shapeProps.Append(new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = cx, Cy = cy }));

            var (preset2, adjustList2) = MapPresetGeometry(shapeType);

            shapeProps.Append(new A.PresetGeometry(adjustList2) { Preset = preset2 });
            wsp.Append(shapeProps);

            var textBodyProps = new Wps.TextBodyProperties() {
                Rotation = 0,
                UseParagraphSpacing = false,
                VerticalOverflow = A.TextVerticalOverflowValues.Overflow,
                HorizontalOverflow = A.TextHorizontalOverflowValues.Overflow,
                Vertical = A.TextVerticalValues.Horizontal,
                Wrap = A.TextWrappingValues.Square,
                LeftInset = 91440,
                TopInset = 45720,
                RightInset = 91440,
                BottomInset = 45720,
                ColumnCount = 1,
                ColumnSpacing = 0,
                RightToLeftColumns = false,
                FromWordArt = false,
                Anchor = A.TextAnchoringTypeValues.Center,
                AnchorCenter = false,
                ForceAntiAlias = false,
                CompatibleLineSpacing = true
            };
            textBodyProps.Append(new A.NoAutoFit());
            wsp.Append(textBodyProps);

            graphicData.Append(wsp);
            graphic.Append(graphicData);
            anchor.Append(graphic);

            var drawing = new Drawing(anchor);
            run.Append(drawing);
            return new WordShape(paragraph._document!, paragraph._paragraph!, run, drawing);
        }

        private static string NormalizeHexNoHash(string? v) {
            if (string.IsNullOrEmpty(v)) return string.Empty;
            var s = v!.Trim();
            if (s.StartsWith("#", StringComparison.Ordinal)) s = s.Substring(1);
            return s.ToLowerInvariant();
        }

        /// <summary>
        /// Gets or sets the fill color as hexadecimal string (no leading '#', lowercase).
        /// </summary>
        public string FillColorHex {
            get {
                if (_rectangle?.FillColor?.Value is string rect) return NormalizeHexNoHash(rect);
                if (_roundRectangle?.FillColor?.Value is string round) return NormalizeHexNoHash(round);
                if (_ellipse?.FillColor?.Value is string ellipse) return NormalizeHexNoHash(ellipse);
                if (_polygon?.FillColor?.Value is string poly) return NormalizeHexNoHash(poly);
                if (_shape?.FillColor?.Value is string shape) return NormalizeHexNoHash(shape);
                if (_wpsShape != null) {
                    var spPr = _wpsShape.GetFirstChild<Wps.ShapeProperties>();
                    var solid = spPr?.GetFirstChild<A.SolidFill>();
                    var rgb = solid?.GetFirstChild<A.RgbColorModelHex>();
                    if (rgb?.Val != null) return NormalizeHexNoHash(rgb.Val.Value);
                }
                return string.Empty;
            }
            set {
                string? v = value;
                if (!string.IsNullOrEmpty(v) && !v.StartsWith("#", StringComparison.Ordinal)) v = "#" + v;
                if (_rectangle != null) _rectangle.FillColor = v;
                if (_roundRectangle != null) _roundRectangle.FillColor = v;
                if (_ellipse != null) _ellipse.FillColor = v;
                if (_polygon != null) _polygon.FillColor = v;
                if (_shape != null) _shape.FillColor = v;
                if (_wpsShape != null && !string.IsNullOrEmpty(v)) {
                    var spPr = _wpsShape.GetFirstChild<Wps.ShapeProperties>();
                    if (spPr != null) {
                        // Remove NoFill if present
                        var noFill = spPr.GetFirstChild<A.NoFill>();
                        noFill?.Remove();

                        var solid = spPr.GetFirstChild<A.SolidFill>();
                        if (solid == null) {
                            solid = new A.SolidFill();
                            // Insert after geometry if possible
                            var geom = (OpenXmlElement?)spPr.GetFirstChild<A.CustomGeometry>() ?? spPr.GetFirstChild<A.PresetGeometry>();
                            if (geom != null) spPr.InsertAfter(solid, geom);
                            else spPr.Append(solid);
                        }
                        var rgb = solid.GetFirstChild<A.RgbColorModelHex>();
                        if (rgb == null) {
                            rgb = new A.RgbColorModelHex();
                            solid.RemoveAllChildren();
                            solid.Append(rgb);
                        }
                        rgb.Val = v!.TrimStart('#');
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the fill color using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public SixLabors.ImageSharp.Color FillColor {
            get {
                var hex = FillColorHex;
                if (string.IsNullOrEmpty(hex)) return SixLabors.ImageSharp.Color.Transparent;
                if (!hex.StartsWith("#", StringComparison.Ordinal)) hex = "#" + hex;
                return SixLabors.ImageSharp.Color.Parse(hex);
            }
            set => FillColorHex = value.ToHexColor();
        }

        /// <summary>
        /// Identifier of the shape.
        /// </summary>
        public string Id {
            get {
                if (_rectangle != null) return _rectangle.Id?.Value ?? string.Empty;
                if (_roundRectangle != null) return _roundRectangle.Id?.Value ?? string.Empty;
                if (_ellipse != null) return _ellipse.Id?.Value ?? string.Empty;
                if (_polygon != null) return _polygon.Id?.Value ?? string.Empty;
                if (_line != null) return _line.Id?.Value ?? string.Empty;
                if (_shape != null) return _shape.Id?.Value ?? string.Empty;
                return string.Empty;
            }
            set {
                if (_rectangle != null) _rectangle.Id = value;
                if (_roundRectangle != null) _roundRectangle.Id = value;
                if (_ellipse != null) _ellipse.Id = value;
                if (_polygon != null) _polygon.Id = value;
                if (_line != null) _line.Id = value;
                if (_shape != null) _shape.Id = value;
            }
        }

        /// <summary>
        /// Optional title of the shape.
        /// </summary>
        public string? Title {
            get {
                if (_rectangle != null) return _rectangle.Title?.Value;
                if (_roundRectangle != null) return _roundRectangle.Title?.Value;
                if (_ellipse != null) return _ellipse.Title?.Value;
                if (_polygon != null) return _polygon.Title?.Value;
                if (_line != null) return _line.Title?.Value;
                if (_shape != null) return _shape.Title?.Value;
                return null;
            }
            set {
                if (_rectangle != null) _rectangle.Title = value;
                if (_roundRectangle != null) _roundRectangle.Title = value;
                if (_ellipse != null) _ellipse.Title = value;
                if (_polygon != null) _polygon.Title = value;
                if (_line != null) _line.Title = value;
                if (_shape != null) _shape.Title = value;
            }
        }

        /// <summary>
        /// Alternative text description of the shape.
        /// </summary>
        public string? Description {
            get {
                if (_rectangle != null) return _rectangle.Alternate?.Value;
                if (_roundRectangle != null) return _roundRectangle.Alternate?.Value;
                if (_ellipse != null) return _ellipse.Alternate?.Value;
                if (_polygon != null) return _polygon.Alternate?.Value;
                if (_line != null) return _line.Alternate?.Value;
                if (_shape != null) return _shape.Alternate?.Value;
                return null;
            }
            set {
                if (_rectangle != null) _rectangle.Alternate = value;
                if (_roundRectangle != null) _roundRectangle.Alternate = value;
                if (_ellipse != null) _ellipse.Alternate = value;
                if (_polygon != null) _polygon.Alternate = value;
                if (_line != null) _line.Alternate = value;
                if (_shape != null) _shape.Alternate = value;
            }
        }

        /// <summary>
        /// Whether the shape is hidden. Stored as "visibility:hidden" in the style string.
        /// </summary>
        public bool? Hidden {
            get {
                var v = GetStyleValue("visibility");
                if (string.IsNullOrEmpty(v)) return null;
                return v == "hidden";
            }
            set {
                if (value == null) {
                    RemoveStyleValue("visibility");
                } else {
                    SetStyleValue("visibility", value.Value ? "hidden" : "visible");
                }
            }
        }

        /// <summary>
        /// Outline color in hex format (no leading '#', lowercase). Null when not applicable.
        /// </summary>
        public string? StrokeColorHex {
            get {
                if (_rectangle != null) return NormalizeHexNoHash(_rectangle.StrokeColor?.Value);
                if (_roundRectangle != null) return NormalizeHexNoHash(_roundRectangle.StrokeColor?.Value);
                if (_ellipse != null) return NormalizeHexNoHash(_ellipse.StrokeColor?.Value);
                if (_polygon != null) return NormalizeHexNoHash(_polygon.StrokeColor?.Value);
                if (_line != null) return NormalizeHexNoHash(_line.StrokeColor?.Value);
                if (_shape != null) return NormalizeHexNoHash(_shape.StrokeColor?.Value);
                if (_wpsShape != null) {
                    var spPr = _wpsShape.GetFirstChild<Wps.ShapeProperties>();
                    var outline = spPr?.GetFirstChild<A.Outline>();
                    var solid = outline?.GetFirstChild<A.SolidFill>();
                    var rgb = solid?.GetFirstChild<A.RgbColorModelHex>();
                    if (rgb?.Val != null) return NormalizeHexNoHash(rgb.Val.Value);
                }
                return null;
            }
            set {
                string? v = value;
                if (!string.IsNullOrEmpty(v) && !v.StartsWith("#", StringComparison.Ordinal)) v = "#" + v;
                if (_rectangle != null) _rectangle.StrokeColor = v;
                if (_roundRectangle != null) _roundRectangle.StrokeColor = v;
                if (_ellipse != null) _ellipse.StrokeColor = v;
                if (_polygon != null) _polygon.StrokeColor = v;
                if (_line != null) _line.StrokeColor = v;
                if (_shape != null) _shape.StrokeColor = v;
                if (_wpsShape != null && !string.IsNullOrEmpty(v)) {
                    var spPr = _wpsShape.GetFirstChild<Wps.ShapeProperties>();
                    if (spPr != null) {
                        var outline = spPr.GetFirstChild<A.Outline>();
                        if (outline == null) {
                            outline = new A.Outline();
                            spPr.Append(outline);
                        }
                        var solid = outline.GetFirstChild<A.SolidFill>();
                        if (solid == null) {
                            solid = new A.SolidFill();
                            outline.RemoveAllChildren<A.FillProperties>();
                            outline.Append(solid);
                        }
                        var rgb = solid.GetFirstChild<A.RgbColorModelHex>();
                        if (rgb == null) {
                            rgb = new A.RgbColorModelHex();
                            solid.RemoveAllChildren();
                            solid.Append(rgb);
                        }
                        rgb.Val = v!.TrimStart('#');
                    }
                }
            }
        }

        /// <summary>
        /// Outline color using <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        public SixLabors.ImageSharp.Color StrokeColor {
            get {
                var hex = StrokeColorHex;
                if (string.IsNullOrEmpty(hex)) return SixLabors.ImageSharp.Color.Transparent;
                if (!hex.StartsWith("#", StringComparison.Ordinal)) hex = "#" + hex;
                return SixLabors.ImageSharp.Color.Parse(hex);
            }
            set => StrokeColorHex = value.ToHexColor();
        }

        /// <summary>
        /// Outline thickness in points.
        /// </summary>
        public double? StrokeWeight {
            get {
                string? v = null;
                if (_rectangle != null) v = _rectangle.StrokeWeight?.Value;
                if (_roundRectangle != null) v ??= _roundRectangle.StrokeWeight?.Value;
                if (_ellipse != null) v ??= _ellipse.StrokeWeight?.Value;
                if (_polygon != null) v ??= _polygon.StrokeWeight?.Value;
                if (_line != null) v ??= _line.StrokeWeight?.Value;
                if (_shape != null) v ??= _shape.StrokeWeight?.Value;
                if (string.IsNullOrEmpty(v)) return null;
                return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
            }
            set {
                string? v = value != null ? $"{value.Value.ToString(CultureInfo.InvariantCulture)}pt" : null;
                if (_rectangle != null) _rectangle.StrokeWeight = v;
                if (_roundRectangle != null) _roundRectangle.StrokeWeight = v;
                if (_ellipse != null) _ellipse.StrokeWeight = v;
                if (_polygon != null) _polygon.StrokeWeight = v;
                if (_line != null) _line.StrokeWeight = v;
                if (_shape != null) _shape.StrokeWeight = v;
                if (_wpsShape != null && value != null) {
                    var spPr = _wpsShape.GetFirstChild<Wps.ShapeProperties>();
                    if (spPr != null) {
                        var outline = spPr.GetFirstChild<A.Outline>();
                        if (outline == null) {
                            outline = new A.Outline();
                            spPr.Append(outline);
                        }
                        outline.Width = (Int32Value)(int)Math.Round(value.Value * EmusPerPoint);
                    }
                }
            }
        }

        /// <summary>
        /// Corner roundness as a fraction between 0 and 1 for rounded rectangles.
        /// </summary>
        public double? ArcSize {
            get {
                var arc = _roundRectangle?.ArcSize;
                if (arc == null) return null;
                var val = arc.Value?.TrimEnd('f');
                if (val == null || !double.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out var num)) return null;
                return num / 65536d;
            }
            set {
                if (_roundRectangle == null || value == null) return;
                var v = (int)Math.Round(value.Value * 65536d);
                _roundRectangle.ArcSize = v.ToString(CultureInfo.InvariantCulture) + "f";
            }
        }

        /// <summary>
        /// Determines whether the outline is drawn.
        /// </summary>
        public bool? Stroked {
            get {
                if (_rectangle?.Stroked != null) return _rectangle.Stroked.Value;
                if (_roundRectangle?.Stroked != null) return _roundRectangle.Stroked.Value;
                if (_ellipse?.Stroked != null) return _ellipse.Stroked.Value;
                if (_polygon?.Stroked != null) return _polygon.Stroked.Value;
                if (_line?.Stroked != null) return _line.Stroked.Value;
                if (_shape?.Stroked != null) return _shape.Stroked.Value;
                return null;
            }
            set {
                if (_rectangle != null) _rectangle.Stroked = value;
                if (_roundRectangle != null) _roundRectangle.Stroked = value;
                if (_ellipse != null) _ellipse.Stroked = value;
                if (_polygon != null) _polygon.Stroked = value;
                if (_line != null) _line.Stroked = value;
                if (_shape != null) _shape.Stroked = value;
            }
        }

        private string? GetStyle() {
            return _rectangle?.Style?.Value ??
                   _roundRectangle?.Style?.Value ??
                   _ellipse?.Style?.Value ??
                   _polygon?.Style?.Value ??
                   _line?.Style?.Value ??
                   _shape?.Style?.Value;
        }

        private void SetStyle(string style) {
            if (_rectangle != null) _rectangle.Style = style;
            if (_roundRectangle != null) _roundRectangle.Style = style;
            if (_ellipse != null) _ellipse.Style = style;
            if (_polygon != null) _polygon.Style = style;
            if (_line != null) _line.Style = style;
            if (_shape != null) _shape.Style = style;
        }

        private string? GetStyleValue(string name) {
            var style = GetStyle();
            if (string.IsNullOrEmpty(style)) return null;
            foreach (var part in style.Split(';')) {
                var kv = part.Split(':');
                if (kv.Length == 2 && kv[0] == name) return kv[1];
            }
            return null;
        }

        /// <summary>
        /// Gets or sets the z-index for VML shapes (style "z-index"). DrawingML is not affected.
        /// </summary>
        public int? ZIndex {
            get {
                var v = GetStyleValue("z-index");
                if (string.IsNullOrEmpty(v)) return null;
                if (int.TryParse(v, out var n)) return n;
                return null;
            }
            set {
                if (value == null) {
                    RemoveStyleValue("z-index");
                } else {
                    SetStyleValue("z-index", value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
            }
        }

        private void SetStyleValue(string name, string value) {
            var style = GetStyle() ?? string.Empty;
            var parts = style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            bool updated = false;
            for (int i = 0; i < parts.Count; i++) {
                var kv = parts[i].Split(':');
                if (kv.Length == 2 && kv[0] == name) {
                    parts[i] = $"{name}:{value}";
                    updated = true;
                    break;
                }
            }
            if (!updated) parts.Add($"{name}:{value}");
            SetStyle(string.Join(";", parts));
        }

        private void RemoveStyleValue(string name) {
            var style = GetStyle();
            if (string.IsNullOrEmpty(style)) return;
            var parts = style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            parts.RemoveAll(p => p.Split(':').FirstOrDefault() == name);
            SetStyle(string.Join(";", parts));
        }

        /// <summary>
        /// Width of the shape in points.
        /// </summary>
        public double Width {
            get {
                var v = GetStyleValue("width");
                if (!string.IsNullOrEmpty(v)) return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
                return 0;
            }
            set => SetStyleValue("width", $"{value.ToString(CultureInfo.InvariantCulture)}pt");
        }

        /// <summary>
        /// Height of the shape in points.
        /// </summary>
        public double Height {
            get {
                var v = GetStyleValue("height");
                if (!string.IsNullOrEmpty(v)) return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
                return 0;
            }
            set => SetStyleValue("height", $"{value.ToString(CultureInfo.InvariantCulture)}pt");
        }

        /// <summary>
        /// Left position of the shape in points. Returns null when not explicitly set.
        /// </summary>
        public double? Left {
            get {
                var v = GetStyleValue("margin-left");
                if (string.IsNullOrEmpty(v)) return null;
                return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
            }
            set {
                if (value == null) {
                    RemoveStyleValue("margin-left");
                } else {
                    SetStyleValue("margin-left", $"{value.Value.ToString(CultureInfo.InvariantCulture)}pt");
                    if (GetStyleValue("position") == null) SetStyleValue("position", "absolute");
                }
            }
        }

        /// <summary>
        /// Top position of the shape in points. Returns null when not explicitly set.
        /// </summary>
        public double? Top {
            get {
                var v = GetStyleValue("margin-top");
                if (string.IsNullOrEmpty(v)) return null;
                return double.Parse(v.Replace("pt", string.Empty), CultureInfo.InvariantCulture);
            }
            set {
                if (value == null) {
                    RemoveStyleValue("margin-top");
                } else {
                    SetStyleValue("margin-top", $"{value.Value.ToString(CultureInfo.InvariantCulture)}pt");
                    if (GetStyleValue("position") == null) SetStyleValue("position", "absolute");
                }
            }
        }

        /// <summary>
        /// Rotation of the shape in degrees. Returns null when not set.
        /// </summary>
        public double? Rotation {
            get {
                var v = GetStyleValue("rotation");
                if (string.IsNullOrEmpty(v)) return null;
                return double.Parse(v, CultureInfo.InvariantCulture);
            }
            set {
                if (value == null) {
                    RemoveStyleValue("rotation");
                } else {
                    SetStyleValue("rotation", value.Value.ToString(CultureInfo.InvariantCulture));
                }
            }
        }

        /// <summary>
        /// Removes the shape from the paragraph.
        /// </summary>
        public void Remove() {
            _run?.Remove();
        }
    }
}
