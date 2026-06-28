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

        private const double EnglishMetricUnitsPerInch = 914400;
        private const double PixelsPerInch = 96;

        internal WordDrawing _Image = null!;
        private ImagePart? _imagePart;
        private string? _externalRelationshipId;
        private WordDocument _document = null!;
        private int? _cropTop;
        private int? _cropBottom;
        private int? _cropLeft;
        private int? _cropRight;
        private ImageFillMode _fillMode = ImageFillMode.Stretch;
        private bool? _useLocalDpi;
        private string? _title;
        private bool? _hidden;
        private bool? _preferRelativeResize;
        private bool? _noChangeAspect;
        private bool? _noCrop;
        private bool? _noMove;
        private bool? _noResize;
        private bool? _noRotation;
        private bool? _noSelection;
        private int? _fixedOpacity;
        private string? _alphaInversionColorHex;
        private int? _blackWhiteThreshold;
        private int? _blurRadius;
        private bool? _blurGrow;
        private string? _colorChangeFromHex;
        private string? _colorChangeToHex;
        private string? _colorReplacementHex;
        private string? _duotoneColor1Hex;
        private string? _duotoneColor2Hex;
        private bool? _grayScale;
        private int? _luminanceBrightness;
        private int? _luminanceContrast;
        private int? _tintAmount;
        private int? _tintHue;

        internal V.Shape? _vmlShape;
        internal V.ImageData? _vmlImageData;


        /// <summary>
        /// Get or set the Image's horizontal position.
        /// </summary>
        public HorizontalPosition horizontalPosition {
            get {
                if (_Image.Inline == null) {
                    var anchor = _Image.Anchor ?? throw new InvalidOperationException("Anchor is missing.");
                    return anchor.HorizontalPosition ?? throw new InvalidOperationException("HorizontalPosition is missing.");
                }

                throw new InvalidOperationException("Inline images do not have HorizontalPosition property.");
            }
            set {
                if (_Image.Inline == null) {
                    var anchor = _Image.Anchor ?? throw new InvalidOperationException("Anchor is missing.");
                    anchor.HorizontalPosition = value;
                } else {
                    throw new InvalidOperationException("Inline images do not have HorizontalPosition property.");
                }
            }
        }

        /// <summary>
        /// Get or set the Image's vertical position.
        /// </summary>
        public VerticalPosition verticalPosition {
            get {
                if (_Image.Inline == null) {
                    var anchor = _Image.Anchor ?? throw new InvalidOperationException("Anchor is missing.");
                    return anchor.VerticalPosition ?? throw new InvalidOperationException("VerticalPosition is missing.");
                }

                throw new InvalidOperationException("Inline images do not have VerticalPosition property.");
            }
            set {
                if (_Image.Inline == null) {
                    var anchor = _Image.Anchor ?? throw new InvalidOperationException("Anchor is missing.");
                    anchor.VerticalPosition = value;
                } else {
                    throw new InvalidOperationException("Inline images do not have VerticalPosition property.");
                }
            }
        }

        /// <summary>
        /// Gets or sets the compression quality for embedded images.
        /// </summary>
        public BlipCompressionValues? CompressionQuality {
            get {
                var picture = GetPicture();
                return picture?.BlipFill?.Blip?.CompressionState?.Value;
            }
            set {
                var picture = GetPicture();
                if (picture?.BlipFill?.Blip != null) {
                    if (value.HasValue) {
                        picture.BlipFill.Blip.CompressionState = value.Value;
                    } else {
                        picture.BlipFill.Blip.CompressionState = null;
                    }
                    SetPicture(picture);
                }
            }
        }

        /// <summary>
        /// Gets the relationship id of the embedded image.
        /// </summary>
        public string? RelationshipId => GetBlip()?.Embed?.Value;

        /// <summary>
        /// Gets the relationship id of an externally linked image, if any.
        /// </summary>
        public string? ExternalRelationshipId => _externalRelationshipId;

        /// <summary>
        /// Indicates whether the image is stored outside the package.
        /// </summary>
        public bool IsExternal => _externalRelationshipId != null;

        /// <summary>
        /// Gets the content type of the embedded image part, or an empty string for external images.
        /// </summary>
        public string ContentType => _imagePart?.ContentType ?? string.Empty;

        /// <summary>
        /// Gets the URI of the externally linked image.
        /// </summary>
        public Uri? ExternalUri {
            get {
                if (_externalRelationshipId == null) return null;
                var part = GetContainingPart();
                var rel = part.ExternalRelationships.FirstOrDefault(r => r.Id == _externalRelationshipId);
                return rel?.Uri;
            }
        }

        /// <summary>
        /// Gets or sets the file path or name for the image.
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// Get or sets the image's file name
        /// </summary>
        public string? FileName {
            get => GetPicture()?.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name;
            set {
                if (value == null) throw new ArgumentNullException(nameof(value));
                var drawingProperties = GetPicture()?.NonVisualPictureProperties?.NonVisualDrawingProperties;
                if (drawingProperties != null) {
                    drawingProperties.Name = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the image's description.
        /// </summary>
        public string? Description {
            get => GetDocProperties()?.Description;
            set => GetWritableDocProperties()?.Description = value;
        }

        /// <summary>
        /// Gets or sets the image's title.
        /// </summary>
        public string? Title {
            get {
                _title = GetDocProperties()?.Title;
                return _title;
            }
            set {
                _title = value;
                GetWritableDocProperties()?.Title = value;
                var drawingProperties = GetPicture()?.NonVisualPictureProperties?.NonVisualDrawingProperties;
                if (drawingProperties != null) {
                    drawingProperties.Title = value;
                }
            }
        }

        /// <summary>
        /// Specifies whether the picture is hidden.
        /// </summary>
        public bool? Hidden {
            get {
                _hidden = GetDocProperties()?.Hidden?.Value;
                return _hidden;
            }
            set {
                _hidden = value;
                GetWritableDocProperties()?.Hidden = value;
                var drawingProperties = GetPicture()?.NonVisualPictureProperties?.NonVisualDrawingProperties;
                if (drawingProperties != null) {
                    drawingProperties.Hidden = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets Width of an image
        /// </summary>
        public double? Width {
            get {
                var inlineCx = _Image.Inline?.Extent?.Cx?.Value;
                if (inlineCx.HasValue) {
                    return inlineCx.Value / EnglishMetricUnitsPerInch * PixelsPerInch;
                }
                var anchorCx = _Image.Anchor?.Extent?.Cx?.Value;
                if (anchorCx.HasValue) {
                    return anchorCx.Value / EnglishMetricUnitsPerInch * PixelsPerInch;
                }
                return null;
            }
            set {
                if (value == null) throw new ArgumentNullException(nameof(value));
                double emuWidth = value.Value * EnglishMetricUnitsPerInch / PixelsPerInch;
                if (_Image.Inline?.Extent != null) {
                    _Image.Inline.Extent.Cx = (long)emuWidth;
                    var picture = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var ext = picture?.ShapeProperties?.Transform2D?.Extents;
                    if (ext != null) {
                        ext.Cx = (Int64Value)emuWidth;
                    }
                } else {
                    var a = _Image.Anchor;
                    if (a?.Extent != null) {
                        a.Extent.Cx = (long)emuWidth;
                        var anchorGraphic = a.OfType<Graphic>().FirstOrDefault();
                        var picture = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        var ext = picture?.ShapeProperties?.Transform2D?.Extents;
                        if (ext != null) {
                            ext.Cx = (Int64Value)emuWidth;
                        }
                    }
                }
                // _Image.Inline.Extent.Cx = (Int64Value)emuWidth;
                //var picture = _Image.Inline.Graphic.GraphicData.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().First();
            }
        }

        /// <summary>
        /// Gets or sets Height of an image
        /// </summary>
        public double? Height {
            get {
                var inlineCy = _Image.Inline?.Extent?.Cy?.Value;
                if (inlineCy.HasValue) {
                    return inlineCy.Value / EnglishMetricUnitsPerInch * PixelsPerInch;
                }
                var anchorCy = _Image.Anchor?.Extent?.Cy?.Value;
                if (anchorCy.HasValue) {
                    return anchorCy.Value / EnglishMetricUnitsPerInch * PixelsPerInch;
                }
                return null;
            }
            set {
                if (value == null) throw new ArgumentNullException(nameof(value));
                if (_Image.Inline?.Extent != null) {
                    double emuHeight = value.Value * EnglishMetricUnitsPerInch / PixelsPerInch;
                    _Image.Inline.Extent.Cy = (Int64Value)emuHeight;
                    var picture = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var ext = picture?.ShapeProperties?.Transform2D?.Extents;
                    if (ext != null) {
                        ext.Cy = (Int64Value)emuHeight;
                    }
                } else {
                    var a = _Image.Anchor;
                    if (a?.Extent != null) {
                        double emuHeight = value.Value * EnglishMetricUnitsPerInch / PixelsPerInch;
                        a.Extent.Cy = (Int64Value)emuHeight;
                        var anchorGraphic = a.OfType<Graphic>().FirstOrDefault();
                        var picture = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        var ext = picture?.ShapeProperties?.Transform2D?.Extents;
                        if (ext != null) {
                            ext.Cy = (Int64Value)emuHeight;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets the image width in EMUs.
        /// </summary>
        public double? EmuWidth {
            get {
                var inlineCx = _Image.Inline?.Extent?.Cx?.Value;
                if (inlineCx.HasValue) return inlineCx.Value;

                var anchorCx = _Image.Anchor?.Extent?.Cx?.Value;
                if (anchorCx.HasValue) return anchorCx.Value;

                return null;
            }
        }
        /// <summary>
        /// Gets the image height in EMUs.
        /// </summary>
        public double? EmuHeight {
            get {
                var inlineCy = _Image.Inline?.Extent?.Cy?.Value;
                if (inlineCy.HasValue) return inlineCy.Value;

                var anchorCy = _Image.Anchor?.Extent?.Cy?.Value;
                if (anchorCy.HasValue) return anchorCy.Value;

                return null;
            }
        }

        /// <summary>
        /// Gets or sets the shape type used to display the image.
        /// </summary>
        public ShapeTypeValues? Shape {
            get {
                var pictureInline = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                var presetInline = pictureInline?.ShapeProperties?.GetFirstChild<PresetGeometry>()?.Preset?.Value;
                if (presetInline.HasValue) return presetInline.Value;

                var anchorGraphic = _Image.Anchor?.OfType<Graphic>().FirstOrDefault();
                var pictureAnchor = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                var presetAnchor = pictureAnchor?.ShapeProperties?.GetFirstChild<PresetGeometry>()?.Preset?.Value;
                if (presetAnchor.HasValue) return presetAnchor.Value;

                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var geom = picture?.ShapeProperties?.GetFirstChild<PresetGeometry>();
                    if (geom != null) geom.Preset = value;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    var picture = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var geom = picture?.ShapeProperties?.GetFirstChild<PresetGeometry>();
                    if (geom != null) geom.Preset = value;
                }

            }
        }

        /// <summary>
        /// Microsoft Office does not seem to fully support this attribute, and ignores this setting.
        /// More information: http://officeopenxml.com/drwSp-SpPr.php
        /// </summary>
        public BlackWhiteModeValues? BlackWiteMode {
            get {
                var pictureInline = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                var bwInline = pictureInline?.ShapeProperties?.BlackWhiteMode?.Value;
                if (bwInline.HasValue) return bwInline.Value;

                var anchorGraphic = _Image.Anchor?.OfType<Graphic>().FirstOrDefault();
                var pictureAnchor = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                var bwAnchor = pictureAnchor?.ShapeProperties?.BlackWhiteMode?.Value;
                if (bwAnchor.HasValue) return bwAnchor.Value;

                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var sp = picture?.ShapeProperties;
                    if (sp != null) {
                        if (value == null) {
                            // no change
                        } else {
                            sp.BlackWhiteMode ??= new EnumValue<BlackWhiteModeValues>();
                            sp.BlackWhiteMode.Value = value.Value;
                        }
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    var picture = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var sp = picture?.ShapeProperties;
                    if (sp != null) {
                        if (value == null) {
                            // no change
                        } else {
                            sp.BlackWhiteMode ??= new EnumValue<BlackWhiteModeValues>();
                            sp.BlackWhiteMode.Value = value.Value;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Gets or sets whether the image is vertically flipped.
        /// </summary>
        public bool? VerticalFlip {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var vf = picture?.ShapeProperties?.Transform2D?.VerticalFlip?.Value;
                    if (vf.HasValue) return vf.Value;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    var picture = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var vf = picture?.ShapeProperties?.Transform2D?.VerticalFlip?.Value;
                    if (vf.HasValue) return vf.Value;
                }

                return null;
            }
            set {
                if (!value.HasValue) return;
                if (_Image.Inline != null) {
                    var picture = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var sp = picture?.ShapeProperties;
                    if (sp != null) {
                        sp.Transform2D ??= new A.Transform2D();
                        sp.Transform2D.VerticalFlip = value.Value;
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    var picture = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var sp = picture?.ShapeProperties;
                    if (sp != null) {
                        sp.Transform2D ??= new A.Transform2D();
                        sp.Transform2D.VerticalFlip = value.Value;
                    }
                }
            }
        }
        /// <summary>
        /// Gets or sets whether the image is horizontally flipped.
        /// </summary>
        public bool? HorizontalFlip {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var hf = picture?.ShapeProperties?.Transform2D?.HorizontalFlip?.Value;
                    if (hf.HasValue) return hf.Value;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    var picture = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var hf = picture?.ShapeProperties?.Transform2D?.HorizontalFlip?.Value;
                    if (hf.HasValue) return hf.Value;
                }

                return null;
            }
            set {
                if (!value.HasValue) return;
                if (_Image.Inline != null) {
                    var picture = _Image.Inline?.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var sp = picture?.ShapeProperties;
                    if (sp != null) {
                        sp.Transform2D ??= new A.Transform2D();
                        sp.Transform2D.HorizontalFlip = value.Value;
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    var picture = anchorGraphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var sp = picture?.ShapeProperties;
                    if (sp != null) {
                        sp.Transform2D ??= new A.Transform2D();
                        sp.Transform2D.HorizontalFlip = value.Value;
                    }
                }
            }
        }


        /// <summary>
        /// Gets or sets the image rotation in degrees.
        /// </summary>
        public int? Rotation {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    if (picture?.ShapeProperties?.Transform2D?.Rotation != null) {
                        return picture.ShapeProperties.Transform2D.Rotation / 10000;
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic?.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        if (picture?.ShapeProperties?.Transform2D?.Rotation != null) {
                            return picture.ShapeProperties.Transform2D.Rotation / 10000;
                        }
                    }
                }

                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    if (picture != null) {
                        var shape = picture.ShapeProperties;
                        if (shape == null) {
                            shape = new ShapeProperties();
                            picture.ShapeProperties = shape;
                        }

                        var transform = shape.Transform2D;
                        if (transform == null) {
                            transform = new A.Transform2D();
                            shape.Transform2D = transform;
                        }

                        transform.Rotation = value == null ? null : value.Value * 10000;
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic?.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        if (picture != null) {
                            var shape = picture.ShapeProperties;
                            if (shape == null) {
                                shape = new ShapeProperties();
                                picture.ShapeProperties = shape;
                            }

                            var transform = shape.Transform2D;
                            if (transform == null) {
                                transform = new A.Transform2D();
                                shape.Transform2D = transform;
                            }

                            transform.Rotation = value == null ? null : value.Value * 10000;
                        }
                    }
                }
            }
        }

        private DocumentFormat.OpenXml.Drawing.Pictures.Picture? GetPicture() {
            if (_Image.Inline != null) {
                return _Image.Inline.Graphic?.GraphicData?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
            }

            if (_Image.Anchor != null) {
                var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                if (anchorGraphic?.GraphicData != null) {
                    return anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                }
            }

            return null;
        }

        private DocProperties? GetDocProperties() {
            if (_Image.Inline != null) {
                return _Image.Inline.DocProperties;
            }

            return _Image.Anchor?.OfType<DocProperties>().FirstOrDefault();
        }

        private DocProperties? GetWritableDocProperties() {
            if (_Image.Inline != null) {
                _Image.Inline.DocProperties ??= new DocProperties();
                return _Image.Inline.DocProperties;
            }

            return _Image.Anchor?.OfType<DocProperties>().FirstOrDefault();
        }

        private void SetPicture(Pic.Picture picture) {
            if (_Image.Inline != null) {
                var graphicData = _Image.Inline.Graphic?.GraphicData;
                if (graphicData != null) {
                    graphicData.RemoveAllChildren<Pic.Picture>();
                    graphicData.AppendChild(picture);
                }
            } else if (_Image.Anchor != null) {
                var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                if (anchorGraphic?.GraphicData != null) {
                    anchorGraphic.GraphicData.RemoveAllChildren<Pic.Picture>();
                    anchorGraphic.GraphicData.AppendChild(picture);
                }
            }
        }

    }
}
