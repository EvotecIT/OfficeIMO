using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Color = SixLabors.ImageSharp.Color;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Specifies the type of watermark to insert into a Word document.
    /// </summary>
    public enum WordWatermarkStyle {
        /// <summary>
        /// Indicates a text watermark.
        /// </summary>
        Text,

        /// <summary>
        /// Indicates an image watermark.
        /// </summary>
        Image
    }

    /// <summary>
    /// Represents a watermark element within a Word document.
    /// </summary>
    public class WordWatermark : WordElement {
        private WordDocument _document;
        private SdtBlock _sdtBlock;
        private WordHeader? _wordHeader;
        private WordSection? _section;
        //private WordParagraph _wordParagraph;

        /// <summary>
        /// Gets or sets the watermark text.
        /// </summary>
        public string Text {
            get {
                var paragraph = _sdtBlock.SdtContentBlock?.ChildElements.OfType<Paragraph>().FirstOrDefault();
                if (paragraph != null) {
                    var run = paragraph.Descendants().OfType<Run>().FirstOrDefault();
                    if (run != null) {
                        var picture = run.Descendants().OfType<Picture>().FirstOrDefault();
                        if (picture != null) {
                            var shape = picture.Descendants().OfType<Shape>().FirstOrDefault();
                            if (shape != null) {
                                TextPath? textPath = shape.GetFirstChild<V.TextPath>();
                                if (textPath != null) {
                                    return textPath.String?.Value ?? string.Empty;
                                }
                            }
                        }
                    }
                }

                return "";
            }
            set {
                var paragraph = _sdtBlock.SdtContentBlock?.ChildElements.OfType<Paragraph>().FirstOrDefault();
                if (paragraph != null) {
                    var run = paragraph.Descendants().OfType<Run>().FirstOrDefault();
                    if (run != null) {
                        var picture = run.Descendants().OfType<Picture>().FirstOrDefault();
                        if (picture != null) {
                            var shape = picture.Descendants().OfType<Shape>().FirstOrDefault();
                            if (shape != null) {
                                TextPath? textPath = shape.GetFirstChild<V.TextPath>();
                                if (textPath != null) {
                                    textPath.String = value;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Get or set rotation of the watermark.
        /// </summary>
        public int? Rotation {
            get {
                var shape = _shape;
                if (shape != null) {

                    // Style = "position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s1025", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
                    //
                    var style = shape.Style?.Value;
                    if (style != null) {
                        var rotation = style.Split(';').FirstOrDefault(c => c.StartsWith("rotation:", StringComparison.Ordinal));
                        if (rotation != null) {
                            var rotationValue = rotation.Split(':').LastOrDefault();
                            if (rotationValue != null) {
                                return int.Parse(rotationValue);
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                var style = shape?.Style?.Value;
                if (shape?.Style != null && style != null) {
                    var rotation = style.Split(';').FirstOrDefault(c => c.StartsWith("rotation:", StringComparison.Ordinal));
                    if (rotation != null) {
                        var rotationValue = rotation.Split(':').LastOrDefault();
                        if (rotationValue != null) {
                            shape.Style.Value = style.Replace(rotation, "rotation:" + value);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the width of the watermark in points.
        /// </summary>
        public double? Width {
            get {
                var shape = _shape;
                if (shape != null) {
                    var style = shape.Style?.Value;
                    if (style != null) {
                        var width = style.Split(';').FirstOrDefault(c => c.StartsWith("width:", StringComparison.Ordinal));
                        if (width != null) {
                            var widthValue = width.Split(':').LastOrDefault();
                            if (widthValue != null) {
                                string stringValue = widthValue.Replace("pt", "");
                                return double.Parse(stringValue, CultureInfo.InvariantCulture);
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                var style = shape?.Style?.Value;
                if (shape?.Style != null && style != null) {
                    var width = style.Split(';').FirstOrDefault(c => c.StartsWith("width:", StringComparison.Ordinal));
                    if (width != null) {
                        var widthValue = width.Split(':').LastOrDefault();
                        if (widthValue != null) {
                            shape.Style.Value = style.Replace(width, "width:" + value + "pt");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the height of the watermark in points.
        /// </summary>
        public double? Height {
            get {
                var shape = _shape;
                if (shape != null) {
                    var style = shape.Style?.Value;
                    if (style != null) {
                        var height = style.Split(';').FirstOrDefault(c => c.StartsWith("height:", StringComparison.Ordinal));
                        if (height != null) {
                            var heightValue = height.Split(':').LastOrDefault();
                            if (heightValue != null) {
                                string stringValue = heightValue.Replace("pt", "");
                                return double.Parse(stringValue, CultureInfo.InvariantCulture);
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                var style = shape?.Style?.Value;
                if (shape?.Style != null && style != null) {
                    var height = style.Split(';').FirstOrDefault(c => c.StartsWith("height:", StringComparison.Ordinal));
                    if (height != null) {
                        var heightValue = height.Split(':').LastOrDefault();
                        if (heightValue != null) {
                            shape.Style.Value = style.Replace(height, "height:" + value + "pt");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Horizontal offset (margin-left) of the watermark in points.
        /// </summary>
        public double? HorizontalOffset {
            get {
                var shape = _shape;
                if (shape != null) {
                    var style = shape.Style?.Value;
                    if (style != null) {
                        var left = style.Split(';').FirstOrDefault(c => c.StartsWith("margin-left:", StringComparison.Ordinal));
                        if (left != null) {
                            var value = left.Split(':').LastOrDefault();
                            if (value != null) {
                                string stringValue = value.Replace("pt", "");
                                return double.Parse(stringValue, CultureInfo.InvariantCulture);
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                var style = shape?.Style?.Value;
                if (shape?.Style != null && style != null) {
                    var left = style.Split(';').FirstOrDefault(c => c.StartsWith("margin-left:", StringComparison.Ordinal));
                    if (left != null) {
                        shape.Style.Value = style.Replace(left, "margin-left:" + value + "pt");
                    }
                }
            }
        }

        /// <summary>
        /// Vertical offset (margin-top) of the watermark in points.
        /// </summary>
        public double? VerticalOffset {
            get {
                var shape = _shape;
                if (shape != null) {
                    var style = shape.Style?.Value;
                    if (style != null) {
                        var top = style.Split(';').FirstOrDefault(c => c.StartsWith("margin-top:", StringComparison.Ordinal));
                        if (top != null) {
                            var value = top.Split(':').LastOrDefault();
                            if (value != null) {
                                string stringValue = value.Replace("pt", "");
                                return double.Parse(stringValue, CultureInfo.InvariantCulture);
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                var style = shape?.Style?.Value;
                if (shape?.Style != null && style != null) {
                    var top = style.Split(';').FirstOrDefault(c => c.StartsWith("margin-top:", StringComparison.Ordinal));
                    if (top != null) {
                        shape.Style.Value = style.Replace(top, "margin-top:" + value + "pt");
                    }
                }
            }
        }

        /// <summary>
        /// Font family used for text watermark.
        /// </summary>
        public string? FontFamily {
            get {
                var shape = _shape;
                if (shape != null) {
                    var textPath = shape.GetFirstChild<V.TextPath>();
                    if (textPath?.Style?.Value is { } style) {
                        var family = style.Split(';').FirstOrDefault(c => c.StartsWith("font-family:", StringComparison.Ordinal));
                        if (family != null) {
                            var value = family.Split(':').LastOrDefault();
                            if (value != null) {
                                return value.Trim('"');
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                if (shape != null) {
                    var textPath = shape.GetFirstChild<V.TextPath>();
                    if (textPath != null) {
                        var style = textPath.Style?.Value ?? string.Empty;
                        var dict = style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(p => p.Split(':'))
                            .ToDictionary(p => p[0], p => p.Length > 1 ? p[1] : string.Empty);
                        dict["font-family"] = "\"" + value + "\"";
                        textPath.Style ??= new StringValue();
                        textPath.Style.Value = string.Join(";", dict.Select(kvp => $"{kvp.Key}:{kvp.Value}"));
                    }
                }
            }
        }

        /// <summary>
        /// Font size of text watermark in points.
        /// </summary>
        public double? FontSize {
            get {
                var shape = _shape;
                if (shape != null) {
                    var textPath = shape.GetFirstChild<V.TextPath>();
                    if (textPath?.Style?.Value is { } style) {
                        var size = style.Split(';').FirstOrDefault(c => c.StartsWith("font-size:", StringComparison.Ordinal));
                        if (size != null) {
                            var value = size.Split(':').LastOrDefault();
                            if (value != null) {
                                string stringValue = value.Replace("pt", "");
                                return double.Parse(stringValue, CultureInfo.InvariantCulture);
                            }
                        }
                    }
                }
                return null;
            }
            set {
                var shape = _shape;
                if (shape != null) {
                    var textPath = shape.GetFirstChild<V.TextPath>();
                    if (textPath != null) {
                        var style = textPath.Style?.Value ?? string.Empty;
                        var dict = style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(p => p.Split(':'))
                            .ToDictionary(p => p[0], p => p.Length > 1 ? p[1] : string.Empty);
                        dict["font-size"] = value + "pt";
                        textPath.Style ??= new StringValue();
                        textPath.Style.Value = string.Join(";", dict.Select(kvp => $"{kvp.Key}:{kvp.Value}"));
                    }
                }
            }
        }

        /// <summary>
        /// Opacity of the watermark fill. Value should be between 0 and 1.
        /// </summary>
        public double? Opacity {
            get {
                if (_shape?.GetFirstChild<V.Fill>()?.Opacity?.Value is string opacity) {
                    return double.Parse(opacity, CultureInfo.InvariantCulture);
                }
                return null;
            }
            set {
                var fill = _shape?.GetFirstChild<V.Fill>();
                if (fill != null && value != null) {
                    fill.Opacity = value.Value.ToString(CultureInfo.InvariantCulture);
                }
            }
        }

        /// <summary>
        /// Get or Set if watermark is stroked.
        /// </summary>
        public bool Stroked {
            get {
                var shape = _shape;
                if (shape?.Stroked?.Value != null) {
                    return shape.Stroked.Value;
                }
                return false;
            }
            set {
                var shape = _shape;
                if (shape?.Stroked != null) {
                    shape.Stroked.Value = value;
                }
            }
        }

        /// <summary>
        /// Get or Set if watermark is allowed in cell.
        /// </summary>
        public bool? AllowInCell {
            get {
                var shape = _shape;
                if (shape?.AllowInCell?.Value != null) {
                    return shape.AllowInCell.Value;
                }
                return null;
            }
            set {
                var shape = _shape;
                if (shape?.AllowInCell != null && value.HasValue) {
                    shape.AllowInCell.Value = value.Value;
                }
            }
        }

        /// <summary>
        /// Gets or sets color of the watermark.
        /// Accepted formats for setting the color are named colors (for example, "red")
        /// or three- or six-digit hexadecimal strings with or without a leading '#'.
        /// Invalid color values will result in <see cref="ArgumentException"/> being thrown.
        /// </summary>
        public SixLabors.ImageSharp.Color? Color {
            get {
                if (ColorHex == "") {
                    return null;
                }

                return SixLabors.ImageSharp.Color.Parse(ColorHex);

            }
            set {
                if (value != null) {
                    this.ColorHex = value.Value.ToHexColor();
                }
            }
        }

        /// <summary>
        /// Gets or sets the fill color of the watermark.
        /// The value can be a named color (e.g., "red") or a three- or six-digit hex value with
        /// or without the leading '#'. Three-digit hex values are expanded to six digits.
        /// Invalid inputs will throw <see cref="ArgumentException"/>.
        /// </summary>
        public string ColorHex {
            get {
                var shape = _shape;
                if (shape?.FillColor?.Value != null) {
                    var color = shape.FillColor.Value;
                    return color.StartsWith("#", StringComparison.Ordinal) ? color.Substring(1) : color;
                }
                return "";
            }
            set {
                var shape = _shape;
                if (shape?.FillColor != null && value != null) {
                    var normalized = Helpers.NormalizeColor(value);
                    if (normalized == null) {
                        throw new ArgumentException("Color value cannot be null or empty.", nameof(value));
                    }
                    var colorWithHash = "#" + normalized;
                    shape.FillColor.Value = colorWithHash;

                    var fill = shape.GetFirstChild<V.Fill>();
                    if (fill != null) {
                        fill.Color = colorWithHash;
                    }

                    var textPath = shape.GetFirstChild<V.TextPath>();
                    if (textPath != null) {
                        textPath.SetAttribute(new OpenXmlAttribute("fillcolor", string.Empty, colorWithHash));
                        textPath.SetAttribute(new OpenXmlAttribute("strokecolor", string.Empty, colorWithHash));
                    }
                }
            }
        }

        private Shape? _shape {
            get {

                var shape = _picture?.Descendants().OfType<Shape>().FirstOrDefault();
                if (shape != null) {
                    return shape;
                }
                return null;
            }
        }

        private Picture? _picture {
            get {
                if (_sdtBlock?.SdtContentBlock != null) {
                    Paragraph? paragraph = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>().FirstOrDefault();
                    if (paragraph != null) {
                        Run? run = paragraph.Descendants().OfType<Run>().FirstOrDefault();
                        if (run != null) {
                            Picture? picture = run.Descendants().OfType<Picture>().FirstOrDefault();
                            if (picture != null) {
                                return picture;
                            }
                        }
                    }
                }
                return null;
            }
        }

        private SdtContentBlock? _sdtContentBlock {
            get {
                var sdtBlock = _sdtBlock;
                if (sdtBlock != null) {
                    SdtContentBlock? sdtContentBlock = sdtBlock.GetFirstChild<SdtContentBlock>();
                    if (sdtContentBlock != null) {
                        return sdtContentBlock;
                    }
                }
                return null;
            }
        }

        private static SdtBlock GetStyle(WordWatermarkStyle style) {
            switch (style) {
                case WordWatermarkStyle.Text: return TextWatermark;
                case WordWatermarkStyle.Image: return Confidential2;
            }

            throw new ArgumentOutOfRangeException(nameof(style));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordWatermark"/> class for the specified section.
        /// </summary>
        /// <param name="wordDocument">The parent document.</param>
        /// <param name="wordSection">The section to which the watermark should be added.</param>
        /// <param name="style">The type of watermark to add.</param>
        /// <param name="textOrFilePath">Text to use or path to the watermark image.</param>
        /// <param name="horizontalOffset">Horizontal offset in points.</param>
        /// <param name="verticalOffset">Vertical offset in points.</param>
        /// <param name="scale">Scale factor for width and height.</param>
        public WordWatermark(WordDocument wordDocument, WordSection wordSection, WordWatermarkStyle style, string textOrFilePath, double? horizontalOffset = null, double? verticalOffset = null, double scale = 1.0) {
            this._document = wordDocument;
            this._section = wordSection ?? throw new ArgumentNullException(nameof(wordSection));

            if (style == WordWatermarkStyle.Text) {
                this._sdtBlock = GetStyle(style);

                this.Text = textOrFilePath;

                if (this._section.Paragraphs.Count == 0) {
                    var body = this._document._document.Body ?? throw new InvalidOperationException("Document body is missing");
                    body.Append(_sdtBlock);
                } else {
                    var lastParagraph = this._section.Paragraphs.Last();
                    var parent = lastParagraph._paragraph.Parent ?? throw new InvalidOperationException("Paragraph parent is missing");
                    parent.Append(_sdtBlock);
                }
            } else {
                this._sdtBlock = GetStyle(style);

                var fileName = System.IO.Path.GetFileName(textOrFilePath);
                using var imageStream = new System.IO.FileStream(textOrFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);

                if (this._section.Paragraphs.Count == 0) {
                    var body = this._document._document.Body ?? throw new InvalidOperationException("Document body is missing");
                    body.Append(_sdtBlock);
                } else {
                    var lastParagraph = this._section.Paragraphs.Last();
                    var parent = lastParagraph._paragraph.Parent ?? throw new InvalidOperationException("Paragraph parent is missing");
                    parent.Append(_sdtBlock);
                }

                var sdtContentBlock = this._sdtBlock.SdtContentBlock ?? throw new InvalidOperationException("Missing content block in watermark");
                var paragraph = sdtContentBlock.GetFirstChild<Paragraph>() ?? sdtContentBlock.AppendChild(new Paragraph());
                var wordParagraph = new WordParagraph(wordDocument, paragraph);
                var imageLocation = WordImage.AddImageToLocation(wordDocument, wordParagraph, imageStream, fileName);
                SetWatermarkImageData(imageLocation);
            }

            if (horizontalOffset != null) {
                this.HorizontalOffset = horizontalOffset.Value;
            }

            if (verticalOffset != null) {
                this.VerticalOffset = verticalOffset.Value;
            }

            if (Math.Abs(scale - 1.0) > double.Epsilon) {
                if (this.Width != null) this.Width = this.Width.Value * scale;
                if (this.Height != null) this.Height = this.Height.Value * scale;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordWatermark"/> class for the specified header.
        /// </summary>
        /// <param name="wordDocument">The parent document.</param>
        /// <param name="wordSection">The section associated with the header.</param>
        /// <param name="wordHeader">The header where the watermark will be placed.</param>
        /// <param name="style">The type of watermark to add.</param>
        /// <param name="textOrFilePath">Text to use or path to the watermark image.</param>
        /// <param name="horizontalOffset">Horizontal offset in points.</param>
        /// <param name="verticalOffset">Vertical offset in points.</param>
        /// <param name="scale">Scale factor for width and height.</param>
        public WordWatermark(WordDocument wordDocument, WordSection wordSection, WordHeader wordHeader, WordWatermarkStyle style, string textOrFilePath, double? horizontalOffset = null, double? verticalOffset = null, double scale = 1.0) {
            this._document = wordDocument;
            this._section = wordSection ?? throw new ArgumentNullException(nameof(wordSection));

            if (wordHeader == null) {
                // user didn't create headers first, so we do it for the user
                wordDocument.AddHeadersAndFooters();
                var hdrs = wordDocument.Header ?? throw new InvalidOperationException("Headers not initialized after AddHeadersAndFooters().");
                var def = hdrs.Default ?? throw new InvalidOperationException("Default header not initialized after AddHeadersAndFooters().");
                wordHeader = def;
            }
            this._wordHeader = wordHeader;

            if (style == WordWatermarkStyle.Text) {
                this._sdtBlock = GetStyle(style);

                this.Text = textOrFilePath;

                wordHeader._header?.Append(_sdtBlock);
            } else {
                this._sdtBlock = GetStyle(style);

                var fileName = System.IO.Path.GetFileName(textOrFilePath);
                using var imageStream = new System.IO.FileStream(textOrFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);

                wordHeader._header?.Append(_sdtBlock);

                var sdtContentBlock = this._sdtBlock.SdtContentBlock ?? throw new InvalidOperationException("Missing content block in watermark");
                var paragraph = sdtContentBlock.GetFirstChild<Paragraph>() ?? sdtContentBlock.AppendChild(new Paragraph());
                var wordParagraph = new WordParagraph(wordDocument, paragraph);
                var imageLocation = WordImage.AddImageToLocation(wordDocument, wordParagraph, imageStream, fileName);
                SetWatermarkImageData(imageLocation);
            }

            if (horizontalOffset != null) {
                this.HorizontalOffset = horizontalOffset.Value;
            }

            if (verticalOffset != null) {
                this.VerticalOffset = verticalOffset.Value;
            }

            if (Math.Abs(scale - 1.0) > double.Epsilon) {
                if (this.Width != null) this.Width = this.Width.Value * scale;
                if (this.Height != null) this.Height = this.Height.Value * scale;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordWatermark"/> class from an existing structured document tag.
        /// </summary>
        /// <param name="wordDocument">The parent document.</param>
        /// <param name="sdtBlock">The structured document tag block representing the watermark.</param>
        public WordWatermark(WordDocument wordDocument, SdtBlock sdtBlock) {
            _document = wordDocument;
            _sdtBlock = sdtBlock;
        }

        private static SdtBlock TextWatermark {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -78212419 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Watermarks" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "003C040D", RsidRunAdditionDefault = "003C040D", ParagraphId = "7710D5F9", TextId = "47C6A96F" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "7FF27861" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@7,l@8,m@5,21600l@6,21600e" };

                V.Formulas formulas1 = new V.Formulas();
                V.Formula formula1 = new V.Formula() { Equation = "sum #0 0 10800" };
                V.Formula formula2 = new V.Formula() { Equation = "prod #0 2 1" };
                V.Formula formula3 = new V.Formula() { Equation = "sum 21600 0 @1" };
                V.Formula formula4 = new V.Formula() { Equation = "sum 0 0 @2" };
                V.Formula formula5 = new V.Formula() { Equation = "sum 21600 0 @3" };
                V.Formula formula6 = new V.Formula() { Equation = "if @0 @3 0" };
                V.Formula formula7 = new V.Formula() { Equation = "if @0 21600 @1" };
                V.Formula formula8 = new V.Formula() { Equation = "if @0 0 @2" };
                V.Formula formula9 = new V.Formula() { Equation = "if @0 @4 21600" };
                V.Formula formula10 = new V.Formula() { Equation = "mid @5 @6" };
                V.Formula formula11 = new V.Formula() { Equation = "mid @8 @5" };
                V.Formula formula12 = new V.Formula() { Equation = "mid @7 @8" };
                V.Formula formula13 = new V.Formula() { Equation = "mid @6 @7" };
                V.Formula formula14 = new V.Formula() { Equation = "sum @6 0 @5" };

                formulas1.Append(formula1);
                formulas1.Append(formula2);
                formulas1.Append(formula3);
                formulas1.Append(formula4);
                formulas1.Append(formula5);
                formulas1.Append(formula6);
                formulas1.Append(formula7);
                formulas1.Append(formula8);
                formulas1.Append(formula9);
                formulas1.Append(formula10);
                formulas1.Append(formula11);
                formulas1.Append(formula12);
                formulas1.Append(formula13);
                formulas1.Append(formula14);
                V.Path path1 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" };
                V.TextPath textPath1 = new V.TextPath() { On = true, FitShape = true };

                V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
                V.ShapeHandle shapeHandle1 = new V.ShapeHandle() { Position = "#0,bottomRight", XRange = "6629,14971" };

                shapeHandles1.Append(shapeHandle1);
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

                shapetype1.Append(formulas1);
                shapetype1.Append(path1);
                shapetype1.Append(textPath1);
                shapetype1.Append(shapeHandles1);
                shapetype1.Append(lock1);

                V.Shape shape1 = new V.Shape() { Id = "PowerPlusWaterMarkObject357476642", Style = "position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:90;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s1025", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
                V.Fill fill1 = new V.Fill() { Opacity = ".5" };
                V.TextPath textPath2 = new V.TextPath() { Style = "font-family:\"Calibri\";font-size:1pt;", String = "Draft" };
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };


                //V.Shape shape1 = new V.Shape() {Id = "PowerPlusWaterMarkObject357533252", Style = "position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s1028", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136"};
                //V.Fill fill1 = new V.Fill() {Opacity = ".5"};
                //V.TextPath textPath2 = new V.TextPath() {Style = "font-family:\"Calibri\";font-size:1pt", String = "Test"};
                //Wvml.TextWrap textWrap1 = new Wvml.TextWrap() {AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin};

                shape1.Append(fill1);
                shape1.Append(textPath2);
                shape1.Append(textWrap1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;
            }
        }

        private static SdtBlock Confidential2 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 1122028455 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Watermarks" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "003C040D", RsidRunAdditionDefault = "00F42210", ParagraphId = "7710D5F9", TextId = "7F2AA104" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "006695D0" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@7,l@8,m@5,21600l@6,21600e" };

                V.Formulas formulas1 = new V.Formulas();
                V.Formula formula1 = new V.Formula() { Equation = "sum #0 0 10800" };
                V.Formula formula2 = new V.Formula() { Equation = "prod #0 2 1" };
                V.Formula formula3 = new V.Formula() { Equation = "sum 21600 0 @1" };
                V.Formula formula4 = new V.Formula() { Equation = "sum 0 0 @2" };
                V.Formula formula5 = new V.Formula() { Equation = "sum 21600 0 @3" };
                V.Formula formula6 = new V.Formula() { Equation = "if @0 @3 0" };
                V.Formula formula7 = new V.Formula() { Equation = "if @0 21600 @1" };
                V.Formula formula8 = new V.Formula() { Equation = "if @0 0 @2" };
                V.Formula formula9 = new V.Formula() { Equation = "if @0 @4 21600" };
                V.Formula formula10 = new V.Formula() { Equation = "mid @5 @6" };
                V.Formula formula11 = new V.Formula() { Equation = "mid @8 @5" };
                V.Formula formula12 = new V.Formula() { Equation = "mid @7 @8" };
                V.Formula formula13 = new V.Formula() { Equation = "mid @6 @7" };
                V.Formula formula14 = new V.Formula() { Equation = "sum @6 0 @5" };

                formulas1.Append(formula1);
                formulas1.Append(formula2);
                formulas1.Append(formula3);
                formulas1.Append(formula4);
                formulas1.Append(formula5);
                formulas1.Append(formula6);
                formulas1.Append(formula7);
                formulas1.Append(formula8);
                formulas1.Append(formula9);
                formulas1.Append(formula10);
                formulas1.Append(formula11);
                formulas1.Append(formula12);
                formulas1.Append(formula13);
                formulas1.Append(formula14);
                V.Path path1 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" };
                V.TextPath textPath1 = new V.TextPath() { On = true, FitShape = true };

                V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
                V.ShapeHandle shapeHandle1 = new V.ShapeHandle() { Position = "#0,bottomRight", XRange = "6629,14971" };

                shapeHandles1.Append(shapeHandle1);
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

                shapetype1.Append(formulas1);
                shapetype1.Append(path1);
                shapetype1.Append(textPath1);
                shapetype1.Append(shapeHandles1);
                shapetype1.Append(lock1);

                V.Shape shape1 = new V.Shape() { Id = "PowerPlusWaterMarkObject357533252", Style = "position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s1028", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
                V.Fill fill1 = new V.Fill() { Opacity = ".5" };
                V.TextPath textPath2 = new V.TextPath() { Style = "font-family:\"Calibri\";font-size:1pt", String = "Test" };
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

                shape1.Append(fill1);
                shape1.Append(textPath2);
                shape1.Append(textWrap1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;
            }
        }

        private void AddWatermarkImage(WordParagraph wordParagraph, WordImageLocation imageLocation) {
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            Picture picture1 = new Picture() { AnchorId = "6CD641E4" };

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t75", CoordinateSize = "21600,21600", Filled = false, Stroked = false, OptionalNumber = 75, PreferRelative = true, EdgePath = "m@4@5l@4@11@9@11@9@5xe" };
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula() { Equation = "if lineDrawn pixelLineWidth 0" };
            V.Formula formula2 = new V.Formula() { Equation = "sum @0 1 0" };
            V.Formula formula3 = new V.Formula() { Equation = "sum 0 0 @1" };
            V.Formula formula4 = new V.Formula() { Equation = "prod @2 1 2" };
            V.Formula formula5 = new V.Formula() { Equation = "prod @3 21600 pixelWidth" };
            V.Formula formula6 = new V.Formula() { Equation = "prod @3 21600 pixelHeight" };
            V.Formula formula7 = new V.Formula() { Equation = "sum @0 0 1" };
            V.Formula formula8 = new V.Formula() { Equation = "prod @6 1 2" };
            V.Formula formula9 = new V.Formula() { Equation = "prod @7 21600 pixelWidth" };
            V.Formula formula10 = new V.Formula() { Equation = "sum @8 21600 0" };
            V.Formula formula11 = new V.Formula() { Equation = "prod @7 21600 pixelHeight" };
            V.Formula formula12 = new V.Formula() { Equation = "sum @10 21600 0" };

            formulas1.Append(formula1);
            formulas1.Append(formula2);
            formulas1.Append(formula3);
            formulas1.Append(formula4);
            formulas1.Append(formula5);
            formulas1.Append(formula6);
            formulas1.Append(formula7);
            formulas1.Append(formula8);
            formulas1.Append(formula9);
            formulas1.Append(formula10);
            formulas1.Append(formula11);
            formulas1.Append(formula12);
            V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle, AllowExtrusion = false };
            Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = true };

            shapetype1.Append(stroke1);
            shapetype1.Append(formulas1);
            shapetype1.Append(path1);
            shapetype1.Append(lock1);

            V.Shape shape1 = new V.Shape() {
                Id = "WordPictureWatermark" + Guid.NewGuid().ToString("N"),
                Style = "position:absolute;margin-left:0;margin-top:0;width:467.3pt;height:148.7pt;z-index:-251658240;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                OptionalString = "_x0000_s1025",
                AllowInCell = false,
                Type = "#_x0000_t75"
            };

            V.ImageData imageData1 = new V.ImageData() {
                Gain = "19661f",
                BlackLevel = "22938f",
                Title = imageLocation.ImageName,
                RelationshipId = imageLocation.RelationshipId
            };

            var style = ImageShapeStyleHelper.GetStyle(shape1);
            double widthPt = Helpers.ConvertPixelsToPoints(imageLocation.Width);
            double heightPt = Helpers.ConvertPixelsToPoints(imageLocation.Height);
            style["width"] = widthPt + "pt";
            style["height"] = heightPt + "pt";
            ImageShapeStyleHelper.SetStyle(shape1, style);

            shape1.Append(imageData1);

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            run1.Append(runProperties1);
            run1.Append(picture1);

            wordParagraph._paragraphProperties?.Remove();
            wordParagraph._paragraph!.Append(paragraphProperties1);
            wordParagraph._paragraph.Append(run1);
        }

        private void SetWatermarkImageData(WordImageLocation imageLocation) {
            var shape = _sdtBlock.Descendants<V.Shape>().FirstOrDefault();
            if (shape == null) return;

            shape.RemoveAllChildren<V.TextPath>();

            V.ImageData imageData1 = new V.ImageData() {
                Gain = "19661f",
                BlackLevel = "22938f",
                Title = imageLocation.ImageName,
                RelationshipId = imageLocation.RelationshipId
            };

            var style = ImageShapeStyleHelper.GetStyle(shape);
            double widthPt = Helpers.ConvertPixelsToPoints(imageLocation.Width);
            double heightPt = Helpers.ConvertPixelsToPoints(imageLocation.Height);
            style["width"] = widthPt + "pt";
            style["height"] = heightPt + "pt";
            ImageShapeStyleHelper.SetStyle(shape, style);

            shape.Append(imageData1);
        }

        /// <summary>
        /// Removes this watermark from the document by deleting the structured
        /// document tag that stores its content.
        /// </summary>
        public void Remove() {
            _sdtBlock.Remove();
        }
    }
}
