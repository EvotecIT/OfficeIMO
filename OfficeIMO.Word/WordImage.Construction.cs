using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Anchor = DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties;
using V = DocumentFormat.OpenXml.Vml;

#nullable enable annotations
using DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents an image contained in a <see cref="WordDocument"/> and provides
    /// functionality to insert and manipulate pictures.
    /// </summary>
    public partial class WordImage : WordElement {

        /// <summary>
        /// Initializes a new image from a file path.
        /// </summary>
        public WordImage(WordDocument document, WordParagraph paragraph, string filePath, double? width, double? height, WrapTextImage wrapImage = WrapTextImage.InLineWithText, string description = "", ShapeTypeValues? shape = null, BlipCompressionValues? compressionQuality = null) {
            FilePath = filePath;
            var fileName = System.IO.Path.GetFileName(filePath);
            using var imageStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            shape ??= ShapeTypeValues.Rectangle; // Set default value if not provided
            compressionQuality ??= BlipCompressionValues.Print; // Set default value if not provided
            AddImage(document, paragraph, imageStream, fileName, width, height, shape.Value, compressionQuality.Value, description, wrapImage);
        }

        /// <summary>
        /// Initializes a new image from a stream.
        /// </summary>
        public WordImage(WordDocument document, WordParagraph paragraph, Stream imageStream, string fileName, double? width, double? height, WrapTextImage wrapImage = WrapTextImage.InLineWithText, string description = "", ShapeTypeValues? shape = null, BlipCompressionValues? compressionQuality = null) {
            FilePath = fileName;
            shape ??= ShapeTypeValues.Rectangle; // Set default value if not provided
            compressionQuality ??= BlipCompressionValues.Print; // Set default value if not provided
            AddImage(document, paragraph, imageStream, fileName, width, height, shape.Value, compressionQuality.Value, description, wrapImage);
        }

        /// <summary>
        /// Initializes a new image from a base64 string.
        /// </summary>
        public WordImage(WordDocument document, WordParagraph paragraph, string base64String, string fileName, double? width, double? height, WrapTextImage wrapImage = WrapTextImage.InLineWithText, string description = "", ShapeTypeValues? shape = null, BlipCompressionValues? compressionQuality = null) {
            FilePath = fileName;
            shape ??= ShapeTypeValues.Rectangle;
            compressionQuality ??= BlipCompressionValues.Print;
            var bytes = Convert.FromBase64String(base64String);
            using var ms = new MemoryStream(bytes);
            AddImage(document, paragraph, ms, fileName, width, height, shape.Value, compressionQuality.Value, description, wrapImage);
        }

        /// <summary>
        /// Initializes an image linked to an external URI.
        /// </summary>
        public WordImage(WordDocument document, WordParagraph paragraph, Uri externalUri, double width, double height, WrapTextImage wrapImage = WrapTextImage.InLineWithText, string description = "", ShapeTypeValues? shape = null, BlipCompressionValues? compressionQuality = null) {
            FilePath = externalUri.ToString();
            shape ??= ShapeTypeValues.Rectangle;
            compressionQuality ??= BlipCompressionValues.Print;
            AddExternalImage(document, paragraph, externalUri, width, height, shape.Value, compressionQuality.Value, description, wrapImage);
        }

        private Graphic GetGraphic(double emuWidth, double emuHeight, string fileName, string relationshipId, ShapeTypeValues shape, BlipCompressionValues compressionQuality, string description = "", bool external = false) {
            var shapeProperties = new ShapeProperties();
            var transform2D = new Transform2D();
            var newOffset = new Offset() { X = 0L, Y = 0L };
            var extents = new Extents() { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight };
            transform2D.Append(newOffset);
            transform2D.Append(extents);
            var presetGeometry = new PresetGeometry(new AdjustValueList()) { Preset = shape };
            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            var graphic = new Graphic();
            var graphicData = new GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            var nvDrawingProps = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() {
                Id = (UInt32Value)0U,
                Name = fileName,
                Description = description
            };
            if (_title != null) nvDrawingProps.Title = _title;
            if (_hidden != null) nvDrawingProps.Hidden = _hidden;

            var nvPicProps = new Pic.NonVisualPictureDrawingProperties();
            if (_preferRelativeResize != null) nvPicProps.PreferRelativeResize = _preferRelativeResize;
            if (_noChangeAspect != null || _noCrop != null || _noMove != null || _noResize != null || _noRotation != null || _noSelection != null) {
                var locks = new A.PictureLocks();
                if (_noChangeAspect != null) locks.NoChangeAspect = _noChangeAspect;
                if (_noCrop != null) locks.NoCrop = _noCrop;
                if (_noMove != null) locks.NoMove = _noMove;
                if (_noResize != null) locks.NoResize = _noResize;
                if (_noRotation != null) locks.NoRotation = _noRotation;
                if (_noSelection != null) locks.NoSelection = _noSelection;
                nvPicProps.Append(locks);
            }

            var nonVisualPictureProperties = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(nvDrawingProps, nvPicProps);

            var blipFlip = new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill();

            var blip = new Blip() { CompressionState = compressionQuality };
            if (external) {
                blip.Link = relationshipId;
            } else {
                blip.Embed = relationshipId;
            }
            if (_fixedOpacity != null) {
                blip.Append(new AlphaReplace() { Alpha = _fixedOpacity.Value * 1000 });
            }
            if (_alphaInversionColorHex != null) {
                blip.Append(new AlphaInverse(new RgbColorModelHex { Val = _alphaInversionColorHex }));
            }
            if (_blackWhiteThreshold != null) {
                blip.Append(new BiLevel { Threshold = _blackWhiteThreshold.Value * 1000 });
            }
            if (_blurRadius != null || _blurGrow != null) {
                blip.Append(new Blur { Radius = _blurRadius ?? 0, Grow = _blurGrow ?? false });
            }
            if (_colorChangeFromHex != null || _colorChangeToHex != null) {
                var cc = new ColorChange();
                if (_colorChangeFromHex != null) cc.ColorFrom = new ColorFrom(new RgbColorModelHex { Val = _colorChangeFromHex });
                if (_colorChangeToHex != null) cc.ColorTo = new ColorTo(new RgbColorModelHex { Val = _colorChangeToHex });
                blip.Append(cc);
            }
            if (_colorReplacementHex != null) {
                blip.Append(new ColorReplacement(new RgbColorModelHex { Val = _colorReplacementHex }));
            }
            if (_duotoneColor1Hex != null || _duotoneColor2Hex != null) {
                var duo = new Duotone();
                if (_duotoneColor1Hex != null) duo.Append(new RgbColorModelHex { Val = _duotoneColor1Hex });
                if (_duotoneColor2Hex != null) duo.Append(new RgbColorModelHex { Val = _duotoneColor2Hex });
                blip.Append(duo);
            }
            if (_grayScale == true) {
                blip.Append(new Grayscale());
            }
            if (_luminanceBrightness != null || _luminanceContrast != null) {
                blip.Append(new LuminanceEffect {
                    Brightness = _luminanceBrightness != null ? new Int32Value(_luminanceBrightness.Value * 1000) : null,
                    Contrast = _luminanceContrast != null ? new Int32Value(_luminanceContrast.Value * 1000) : null
                });
            }
            if (_tintAmount != null || _tintHue != null) {
                blip.Append(new TintEffect {
                    Amount = _tintAmount != null ? new Int32Value(_tintAmount.Value * 1000) : null,
                    Hue = _tintHue != null ? new Int32Value(_tintHue.Value * 60000) : null
                });
            }

            // https://stackoverflow.com/questions/33521914/value-of-blipextension-schema-uri-28a0092b-c50c-407e-a947-70e740481c1c
            var blipExtensionList = new BlipExtensionList();
            blip.Append(blipExtensionList);

            var blipExtension1 = new BlipExtension() {
                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
            };

            DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi() { Val = _useLocalDpi ?? false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            blipExtension1.Append(useLocalDpi1);

            blipExtensionList.Append(blipExtension1);

            if (System.IO.Path.GetExtension(fileName).Equals(".svg", StringComparison.OrdinalIgnoreCase)) {
                // Add Office 2010 a14:svgBlip extension that points to the SVG ImagePart
                var svgExt = new BlipExtension() {
                    Uri = "{96DAC541-7B7A-43D3-8B79-8F7C33B92B69}"
                };
                // Create unknown element with correct prefix and namespace
                var svgBlip = new OpenXmlUnknownElement(
                    "a14",
                    "svgBlip",
                    "http://schemas.microsoft.com/office/drawing/2010/main");
                svgBlip.SetAttribute(new OpenXmlAttribute(
                    "r",
                    "embed",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    relationshipId));
                svgExt.Append(svgBlip);
                blipExtensionList.Append(svgExt);
            }

            blipFlip.Append(blip);
            if (_cropTop != null || _cropBottom != null || _cropLeft != null || _cropRight != null) {
                var srcRect = new SourceRectangle();
                if (_cropTop != null) srcRect.Top = _cropTop;
                if (_cropBottom != null) srcRect.Bottom = _cropBottom;
                if (_cropLeft != null) srcRect.Left = _cropLeft;
                if (_cropRight != null) srcRect.Right = _cropRight;
                blipFlip.Append(srcRect);
            }

            switch (_fillMode) {
                case ImageFillMode.Stretch:
                    blipFlip.Append(new Stretch(new FillRectangle()));
                    break;
                case ImageFillMode.Tile:
                    blipFlip.AppendChild(new Tile());
                    break;
                case ImageFillMode.Fit:
                    blipFlip.Append(new Stretch());
                    break;
                case ImageFillMode.Center:
                    blipFlip.AppendChild(new Tile { Alignment = RectangleAlignmentValues.Center });
                    break;
            }

            var picture = new DocumentFormat.OpenXml.Drawing.Pictures.Picture();

            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFlip);
            picture.Append(shapeProperties);

            graphic.Append(graphicData);
            graphicData.Append(picture);

            return graphic;
        }

        private Inline GetInline(double emuWidth, double emuHeight, string imageName, string fileName, string relationshipId, ShapeTypeValues shape, BlipCompressionValues compressionQuality, string description = "", bool external = false) {
            var inline = new Inline() {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946"
            };
            inline.Append(new Extent() { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight });
            inline.Append(new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });
            inline.Append(new DocProperties() {
                Id = (UInt32Value)1U,
                Name = imageName,
                Description = description,
                Title = _title,
                Hidden = _hidden
            });
            inline.Append(new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new GraphicFrameLocks() { NoChangeAspect = true }));
            inline.Append(GetGraphic(emuWidth, emuHeight, fileName, relationshipId, shape, compressionQuality, description, external));

            return inline;
        }

        private Anchor GetAnchor(double emuWidth, double emuHeight, Graphic graphic, string imageName, string description, WrapTextImage wrapImage) {
            bool behindDoc;
            if (wrapImage == WrapTextImage.BehindText) {
                behindDoc = true;
            } else if (wrapImage == WrapTextImage.InFrontOfText) {
                behindDoc = false;
            } else {
                // i think this is default for other cases, except for InlineText which is handled outside of Anchor
                behindDoc = true;
            }

            Anchor anchor1 = new Anchor() {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)114300U,
                DistanceFromRight = (UInt32Value)114300U,
                SimplePos = false,
                RelativeHeight = (UInt32Value)251658240U,
                BehindDoc = behindDoc,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true,
                EditId = "554B0A79",
                AnchorId = "7FB8C9EF"
            };

            SimplePosition simplePosition1 = new SimplePosition() { X = 0L, Y = 0L };
            anchor1.Append(simplePosition1);

            HorizontalPosition horizontalPosition1 = new HorizontalPosition() { RelativeFrom = HorizontalRelativePositionValues.Column };
            PositionOffset positionOffset1 = new PositionOffset { Text = "0" };
            horizontalPosition1.Append(positionOffset1);
            anchor1.Append(horizontalPosition1);

            VerticalPosition verticalPosition1 = new VerticalPosition() { RelativeFrom = VerticalRelativePositionValues.Paragraph };
            PositionOffset positionOffset2 = new PositionOffset { Text = "0" };
            verticalPosition1.Append(positionOffset2);
            anchor1.Append(verticalPosition1);

            Extent extent1 = new Extent() { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight };
            anchor1.Append(extent1);

            EffectExtent effectExtent1 = new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 1905L };
            anchor1.Append(effectExtent1);

            WordWrapTextImage.AppendWrapTextImage(anchor1, wrapImage);

            DocProperties docProperties1 = new DocProperties() {
                Id = (UInt32Value)1U,
                Name = imageName,
                Description = description,
                Title = _title,
                Hidden = _hidden
            };
            anchor1.Append(docProperties1);

            var nonVisualGraphicFrameDrawingProperties1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties();
            GraphicFrameLocks graphicFrameLocks1 = new GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);

            anchor1.Append(graphic);

            RelativeWidth relativeWidth1 = new RelativeWidth() { ObjectId = SizeRelativeHorizontallyValues.Page };
            PercentageWidth percentageWidth1 = new PercentageWidth { Text = "0" };
            relativeWidth1.Append(percentageWidth1);
            anchor1.Append(relativeWidth1);

            RelativeHeight relativeHeight1 = new RelativeHeight() { RelativeFrom = SizeRelativeVerticallyValues.Page };
            PercentageHeight percentageHeight1 = new PercentageHeight { Text = "0" };
            relativeHeight1.Append(percentageHeight1);
            anchor1.Append(relativeHeight1);

            return anchor1;
        }

        private Blip? GetBlip() {
            if (_Image.Inline != null) {
                var picture = _Image.Inline.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                return picture?.BlipFill?.Blip;
            } else if (_Image.Anchor != null) {
                var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                if (anchorGraphic?.GraphicData != null) {
                    var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    return picture?.BlipFill?.Blip;
                }
            }
            return null;
        }

    }
}
