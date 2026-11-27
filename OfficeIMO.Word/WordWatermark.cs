using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using V = DocumentFormat.OpenXml.Vml;

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
            get => _textPath?.String?.Value ?? string.Empty;
            set {
                var textPath = _textPath;
                if (textPath != null) {
                    textPath.String = value;
                }
            }
        }

        /// <summary>
        /// Get or set rotation of the watermark.
        /// </summary>
        public int? Rotation {
            get {
                var rotationValue = GetShapeStyleComponent("rotation");
                return rotationValue == null ? null : int.Parse(rotationValue);
            }
            set => SetShapeStyleComponent("rotation", value?.ToString() ?? string.Empty);
        }

        /// <summary>Gets or sets the physical watermark width. Null removes the stored width.</summary>
        public global::OfficeIMO.OfficeLength? Width {
            get => GetShapeLengthComponent("width");
            set => SetShapeLengthComponent("width", value);
        }

        /// <summary>Gets or sets the physical watermark height. Null removes the stored height.</summary>
        public global::OfficeIMO.OfficeLength? Height {
            get => GetShapeLengthComponent("height");
            set => SetShapeLengthComponent("height", value);
        }

        /// <summary>Gets or sets the signed horizontal offset. Null removes the stored margin-left.</summary>
        public global::OfficeIMO.OfficeLength? HorizontalOffset {
            get => GetShapeLengthComponent("margin-left");
            set => SetShapeLengthComponent("margin-left", value);
        }

        /// <summary>Gets or sets the signed vertical offset. Null removes the stored margin-top.</summary>
        public global::OfficeIMO.OfficeLength? VerticalOffset {
            get => GetShapeLengthComponent("margin-top");
            set => SetShapeLengthComponent("margin-top", value);
        }

        /// <summary>
        /// Font family used for text watermark.
        /// </summary>
        public string? FontFamily {
            get {
                var value = GetTextPathStyleComponent("font-family");
                return value?.Trim('"');
            }
            set => SetTextPathStyleComponent("font-family", "\"" + value + "\"");
        }

        /// <summary>
        /// Font size of text watermark in points.
        /// </summary>
        public global::OfficeIMO.OfficeLength? FontSize {
            get => NativeFontSizePoints is double points ? global::OfficeIMO.OfficeLength.Points(points) : null;
            set => NativeFontSizePoints = value?.ToPoints();
        }

        /// <summary>Gets or sets the native VML size without quantizing imported values.</summary>
        internal double? NativeFontSizePoints {
            get {
                var value = GetTextPathStyleComponent("font-size");
                if (value == null) {
                    return null;
                }

                return ParseWatermarkStylePoint(value);
            }
            set => SetTextPathStyleComponent("font-size", value + "pt");
        }

        /// <summary>
        /// Opacity of the watermark fill. Value should be between 0 and 1.
        /// </summary>
        public double? Opacity {
            get {
                if (_shape?.GetFirstChild<V.Fill>()?.Opacity?.Value is string opacity) {
                    return ParseOpacity(opacity);
                }
                return null;
            }
            set {
                var fill = _shape?.GetFirstChild<V.Fill>();
                if (fill != null && value != null) {
                    fill.Opacity = OfficeIMO.Core.Internal.OfficeRoundTripText.FormatPositional(value.Value);
                }
            }
        }

        private static double? ParseOpacity(string? opacity) {
            if (string.IsNullOrWhiteSpace(opacity)) {
                return null;
            }

            string value = opacity!.Trim();
            double? parsed;
            if (value.EndsWith("%", StringComparison.OrdinalIgnoreCase)) {
                parsed = TryParseDouble(value.Substring(0, value.Length - 1), out double percent) ? percent / 100D : null;
            } else if (value.EndsWith("f", StringComparison.OrdinalIgnoreCase) &&
                       TryParseDouble(value.Substring(0, value.Length - 1), out double fixedPoint)) {
                parsed = fixedPoint / 65536D;
            } else {
                parsed = TryParseDouble(value, out double numeric) ? numeric : null;
            }

            return parsed.HasValue ? Math.Max(0D, Math.Min(1D, parsed.Value)) : null;
        }

        private static bool TryParseDouble(string value, out double result) =>
            OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out result) ||
            OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value.Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out result);

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
        public OfficeIMO.OfficeColor? Color {
            get {
                string value = ColorValue;
                if (value == "") {
                    return null;
                }

                return OfficeIMO.OfficeColor.Parse(value);

            }
            set {
                if (value != null) {
                    this.ColorValue = value.Value.ToRgbHex();
                }
            }
        }

        /// <summary>
        /// Native watermark fill colour without the leading '#'.
        /// </summary>
        internal string ColorValue {
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
            get => _picture?.Descendants().OfType<Shape>().FirstOrDefault();
        }

        private Picture? _picture {
            get => _sdtBlock?.SdtContentBlock?
                .ChildElements.OfType<Paragraph>().FirstOrDefault()?
                .Descendants().OfType<Run>().FirstOrDefault()?
                .Descendants().OfType<Picture>().FirstOrDefault();
        }

        private SdtContentBlock? _sdtContentBlock {
            get => _sdtBlock?.GetFirstChild<SdtContentBlock>();
        }

        private V.TextPath? _textPath => _shape?.GetFirstChild<V.TextPath>();

        internal bool HasImage {
            get {
                string? relationshipId = GetImageRelationshipId();
                return !string.IsNullOrWhiteSpace(relationshipId);
            }
        }

        internal bool TryGetImageBytes(out byte[] bytes, out string? unsupportedReason,
            System.Threading.CancellationToken cancellationToken = default) {
            bytes = Array.Empty<byte>();
            unsupportedReason = null;
            cancellationToken.ThrowIfCancellationRequested();

            string? relationshipId = GetImageRelationshipId();
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                unsupportedReason = "Watermark image relationship is missing.";
                return false;
            }

            if (!TryGetImagePart(relationshipId!, out ImagePart? imagePart) || imagePart == null) {
                unsupportedReason = "Watermark image relationship '" + relationshipId + "' could not be resolved.";
                return false;
            }

            using Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            bytes = OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(stream, cancellationToken);
            if (bytes.Length == 0) {
                unsupportedReason = "Watermark image part is empty.";
                return false;
            }

            return true;
        }

        private string? GetImageRelationshipId() =>
            _shape?.GetFirstChild<V.ImageData>()?.RelationshipId?.Value;

        private bool TryGetImagePart(string relationshipId, out ImagePart? imagePart) {
            imagePart = null;

            HeaderPart? containingHeaderPart = _shape?.Ancestors<Header>().FirstOrDefault()?.HeaderPart ??
                                               _wordHeader?._header?.HeaderPart;
            if (TryGetImagePart(containingHeaderPart, relationshipId, out imagePart)) {
                return true;
            }

            FooterPart? containingFooterPart = _shape?.Ancestors<Footer>().FirstOrDefault()?.FooterPart;
            if (TryGetImagePart(containingFooterPart, relationshipId, out imagePart)) {
                return true;
            }

            MainDocumentPart? mainPart = _document?._wordprocessingDocument?.MainDocumentPart;
            if (TryGetImagePart(mainPart, relationshipId, out imagePart)) {
                return true;
            }

            if (mainPart != null) {
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    if (TryGetImagePart(headerPart, relationshipId, out imagePart)) {
                        return true;
                    }
                }

                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    if (TryGetImagePart(footerPart, relationshipId, out imagePart)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryGetImagePart(OpenXmlPartContainer? container, string relationshipId, out ImagePart? imagePart) {
            imagePart = null;
            if (container == null) {
                return false;
            }

            try {
                imagePart = container.GetPartById(relationshipId) as ImagePart;
                return imagePart != null;
            } catch (ArgumentOutOfRangeException) {
                return false;
            }
        }

        /// <summary>Converts a native VML point value to a whole-EMU public length.</summary>
        private global::OfficeIMO.OfficeLength? GetShapeLengthComponent(string key) {
            string? value = GetShapeStyleComponent(key);
            double? points = value == null ? null : ParseWatermarkStylePoint(value);
            return points.HasValue ? global::OfficeIMO.OfficeLength.Points(points.Value) : null;
        }

        /// <summary>Writes invariant VML points, or removes the component when unset.</summary>
        private void SetShapeLengthComponent(string key, global::OfficeIMO.OfficeLength? value) {
            var shape = _shape;
            if (shape == null) return;

            var style = ImageShapeStyleHelper.GetStyle(shape);
            if (value.HasValue) {
                style[key] = OfficeIMO.Core.Internal.OfficeRoundTripText.Format(value.Value.ToPoints(), CultureInfo.InvariantCulture) + "pt";
            } else {
                style.Remove(key);
            }
            ImageShapeStyleHelper.SetStyle(shape, style);
        }

        private string? GetShapeStyleComponent(string key) =>
            GetStyleComponent(_shape?.Style?.Value, key);

        private void SetShapeStyleComponent(string key, string replacementValue) {
            var shape = _shape;
            var style = shape?.Style?.Value;
            if (shape?.Style == null || style == null) {
                return;
            }

            var component = GetStyleSegment(style, key);
            if (component == null) {
                return;
            }

            shape.Style.Value = style.Replace(component, key + ":" + replacementValue);
        }

        private string? GetTextPathStyleComponent(string key) =>
            GetStyleComponent(_textPath?.Style?.Value, key);

        private void SetTextPathStyleComponent(string key, string replacementValue) {
            var textPath = _textPath;
            if (textPath == null) {
                return;
            }

            var dict = ParseStyleMap(textPath.Style?.Value ?? string.Empty);
            dict[key] = replacementValue;
            textPath.Style ??= new StringValue();
            textPath.Style.Value = string.Join(";", dict.Select(kvp => $"{kvp.Key}:{kvp.Value}"));
        }

        private static string? GetStyleComponent(string? style, string key) {
            var component = GetStyleSegment(style, key);
            return component?.Split(':').LastOrDefault();
        }

        private static string? GetStyleSegment(string? style, string key) {
            if (style == null) {
                return null;
            }

            return style.Split(';').FirstOrDefault(component => component.StartsWith(key + ":", StringComparison.Ordinal));
        }

        private static Dictionary<string, string> ParseStyleMap(string style) =>
            style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(part => part.Split(':'))
                .ToDictionary(part => part[0], part => part.Length > 1 ? part[1] : string.Empty);

        private static double? ParseWatermarkStylePoint(string value) {
            string stringValue = value.Trim();
            if (stringValue.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) {
                stringValue = stringValue.Substring(0, stringValue.Length - 2);
                return ParseWatermarkStyleNumber(stringValue);
            }

            if (stringValue.EndsWith("in", StringComparison.OrdinalIgnoreCase)) {
                return ParseWatermarkStyleNumber(stringValue.Substring(0, stringValue.Length - 2)) * 72D;
            }

            if (stringValue.EndsWith("cm", StringComparison.OrdinalIgnoreCase)) {
                return ParseWatermarkStyleNumber(stringValue.Substring(0, stringValue.Length - 2)) * 28.3464566929D;
            }

            if (stringValue.EndsWith("mm", StringComparison.OrdinalIgnoreCase)) {
                return ParseWatermarkStyleNumber(stringValue.Substring(0, stringValue.Length - 2)) * 2.83464566929D;
            }

            if (stringValue.EndsWith("px", StringComparison.OrdinalIgnoreCase)) {
                return ParseWatermarkStyleNumber(stringValue.Substring(0, stringValue.Length - 2)) * 0.75D;
            }

            return ParseWatermarkStyleNumber(stringValue);
        }

        private static double? ParseWatermarkStyleNumber(string stringValue) {
            if (OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(stringValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double result)) {
                return result;
            }

            stringValue = stringValue.Replace(',', '.');
            return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(stringValue, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
                ? result
                : null;
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="WordWatermark"/> class for the specified section.
        /// </summary>
        /// <param name="wordDocument">The parent document.</param>
        /// <param name="wordSection">The section to which the watermark should be added.</param>
        /// <param name="style">The type of watermark to add.</param>
        /// <param name="textOrFilePath">Text to use or path to the watermark image.</param>
        /// <param name="horizontalOffset">Optional signed horizontal offset.</param>
        /// <param name="verticalOffset">Optional signed vertical offset.</param>
        /// <param name="scale">Scale factor for width and height.</param>
        public WordWatermark(WordDocument wordDocument, WordSection wordSection, WordWatermarkStyle style, string textOrFilePath, global::OfficeIMO.OfficeLength? horizontalOffset = null, global::OfficeIMO.OfficeLength? verticalOffset = null, double scale = 1.0) {
            wordDocument.EnsureStyleDefinition("Header");
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
        /// <param name="horizontalOffset">Optional signed horizontal offset.</param>
        /// <param name="verticalOffset">Optional signed vertical offset.</param>
        /// <param name="scale">Scale factor for width and height.</param>
        public WordWatermark(WordDocument wordDocument, WordSection wordSection, WordHeader wordHeader, WordWatermarkStyle style, string textOrFilePath, global::OfficeIMO.OfficeLength? horizontalOffset = null, global::OfficeIMO.OfficeLength? verticalOffset = null, double scale = 1.0) {
            wordDocument.EnsureStyleDefinition("Header");
            this._document = wordDocument;
            this._section = wordSection ?? throw new ArgumentNullException(nameof(wordSection));

            if (wordHeader == null) {
                // user didn't create headers first, so we do it for the user
                wordHeader = wordDocument.PageSetup.Header;
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
        internal WordWatermark(WordDocument wordDocument, SdtBlock sdtBlock) {
            _document = wordDocument;
            _sdtBlock = sdtBlock;
        }

        private SdtBlock GetStyle(WordWatermarkStyle style) {
            if (_document == null) {
                throw new InvalidOperationException("Document context is not initialized for watermark creation.");
            }

            // The image watermark starts from the same picture template, "Image" in the watermark gallery.
            var block = WordEmbeddedMarkup.CreateBuildingBlock(WordEmbeddedMarkup.Watermarks, style.ToString(), nameof(style));
            _document.AssignNewSdtIds(block);
            return block;
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
            style["width"] = OfficeIMO.Core.Internal.OfficeRoundTripText.Format(widthPt, CultureInfo.InvariantCulture) + "pt";
            style["height"] = OfficeIMO.Core.Internal.OfficeRoundTripText.Format(heightPt, CultureInfo.InvariantCulture) + "pt";
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
            style["width"] = OfficeIMO.Core.Internal.OfficeRoundTripText.Format(widthPt, CultureInfo.InvariantCulture) + "pt";
            style["height"] = OfficeIMO.Core.Internal.OfficeRoundTripText.Format(heightPt, CultureInfo.InvariantCulture) + "pt";
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
