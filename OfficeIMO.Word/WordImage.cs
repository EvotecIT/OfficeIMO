using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using Anchor = DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties;

#nullable enable annotations
using DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents an image contained in a <see cref="WordDocument"/> and provides
    /// functionality to insert and manipulate pictures.
    /// </summary>
    public class WordImage : WordElement {
        private const double EnglishMetricUnitsPerInch = 914400;
        private const double PixelsPerInch = 96;

        internal Drawing _Image;
        private ImagePart _imagePart;
        private string _externalRelationshipId;
        private WordDocument _document;
        private int? _cropTop;
        private int? _cropBottom;
        private int? _cropLeft;
        private int? _cropRight;
        private ImageFillMode _fillMode = ImageFillMode.Stretch;
        private bool? _useLocalDpi;
        private string _title;
        private bool? _hidden;
        private bool? _preferRelativeResize;
        private bool? _noChangeAspect;
        private bool? _noCrop;
        private bool? _noMove;
        private bool? _noResize;
        private bool? _noRotation;
        private bool? _noSelection;
        private int? _fixedOpacity;
        private string _alphaInversionColorHex;
        private int? _blackWhiteThreshold;
        private int? _blurRadius;
        private bool? _blurGrow;
        private string _colorChangeFromHex;
        private string _colorChangeToHex;
        private string _colorReplacementHex;
        private string _duotoneColor1Hex;
        private string _duotoneColor2Hex;
        private bool? _grayScale;
        private int? _luminanceBrightness;
        private int? _luminanceContrast;
        private int? _tintAmount;
        private int? _tintHue;

        /// <summary>
        /// Get or set the Image's horizontal position.
        /// </summary>
        public HorizontalPosition horizontalPosition {
            get {
                if (_Image.Inline == null) {
                    var anchor = _Image.Anchor;
                    var hPosition = anchor.HorizontalPosition;
                    return hPosition;
                } else {
                    throw new InvalidOperationException("Inline images do not have HorizontalPosition property.");
                }
            }
            set {
                if (_Image.Inline == null) {
                    var anchor = _Image.Anchor;
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
                    var anchor = _Image.Anchor;
                    var vPosition = anchor.VerticalPosition;
                    return vPosition;
                } else {
                    throw new InvalidOperationException("Inline images do not have VerticalPosition property.");
                }
            }
            set {
                if (_Image.Inline == null) {
                    var anchor = _Image.Anchor;
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
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    return picture.BlipFill.Blip.CompressionState;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        return picture.BlipFill.Blip.CompressionState;
                    }
                }
                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    if (picture != null) {
                        if (picture.BlipFill != null) {
                            if (picture.BlipFill.Blip != null) {
                                picture.BlipFill.Blip.CompressionState = value;
                            }
                        }
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        if (picture != null) {
                            if (picture.BlipFill != null) {
                                if (picture.BlipFill.Blip != null) {
                                    picture.BlipFill.Blip.CompressionState = value;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets the relationship id of the embedded image.
        /// </summary>
        public string RelationshipId {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    return picture.BlipFill.Blip.Embed;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        return picture.BlipFill.Blip.Embed;
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the relationship id of an externally linked image, if any.
        /// </summary>
        public string ExternalRelationshipId => _externalRelationshipId;

        /// <summary>
        /// Indicates whether the image is stored outside the package.
        /// </summary>
        public bool IsExternal => _externalRelationshipId != null;

        /// <summary>
        /// Gets the URI of the externally linked image.
        /// </summary>
        public Uri ExternalUri {
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
        public string FilePath { get; set; }

        /// <summary>
        /// Get or sets the image's file name
        /// </summary>
        public string FileName {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    return picture.NonVisualPictureProperties.NonVisualDrawingProperties.Name;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        return picture.NonVisualPictureProperties.NonVisualDrawingProperties.Name;
                    }
                }
                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    picture.NonVisualPictureProperties.NonVisualDrawingProperties.Name = value;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        picture.NonVisualPictureProperties.NonVisualDrawingProperties.Name = value;
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the image's description.
        /// </summary>
        public string Description {
            get {

                if (_Image.Inline != null) {
                    return _Image.Inline.DocProperties.Description;

                } else if (_Image.Anchor != null) {
                    var anchoDocPropertiesr = _Image.Anchor.OfType<DocProperties>().FirstOrDefault();
                    return anchoDocPropertiesr.Description;
                }

                return null;
            }
            set {
                if (_Image.Inline != null) {
                    _Image.Inline.DocProperties.Description = value;
                } else if (_Image.Anchor != null) {
                    var anchoDocPropertiesr = _Image.Anchor.OfType<DocProperties>().FirstOrDefault();
                    anchoDocPropertiesr.Description = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the image's title.
        /// </summary>
        public string Title {
            get {
                if (_Image.Inline != null) {
                    _title = _Image.Inline.DocProperties.Title;
                } else if (_Image.Anchor != null) {
                    var docProp = _Image.Anchor.OfType<DocProperties>().FirstOrDefault();
                    _title = docProp?.Title;
                }
                return _title;
            }
            set {
                _title = value;
                if (_Image.Inline != null) {
                    _Image.Inline.DocProperties.Title = value;
                } else if (_Image.Anchor != null) {
                    var docProp = _Image.Anchor.OfType<DocProperties>().FirstOrDefault();
                    if (docProp != null) docProp.Title = value;
                }
                var pic = GetPicture();
                if (pic != null) {
                    var nv = pic.NonVisualPictureProperties;
                    if (nv != null) {
                        nv.NonVisualDrawingProperties.Title = value;
                    }
                }
            }
        }

        /// <summary>
        /// Specifies whether the picture is hidden.
        /// </summary>
        public bool? Hidden {
            get {
                if (_Image.Inline != null) {
                    _hidden = _Image.Inline.DocProperties.Hidden?.Value;
                } else if (_Image.Anchor != null) {
                    var docProp = _Image.Anchor.OfType<DocProperties>().FirstOrDefault();
                    _hidden = docProp?.Hidden?.Value;
                }
                return _hidden;
            }
            set {
                _hidden = value;
                if (_Image.Inline != null) {
                    _Image.Inline.DocProperties.Hidden = value;
                } else if (_Image.Anchor != null) {
                    var docProp = _Image.Anchor.OfType<DocProperties>().FirstOrDefault();
                    if (docProp != null) docProp.Hidden = value;
                }
                var pic = GetPicture();
                if (pic != null) {
                    var nv = pic.NonVisualPictureProperties;
                    if (nv != null) {
                        nv.NonVisualDrawingProperties.Hidden = value;
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets Width of an image
        /// </summary>
        public double? Width {
            get {
                if (_Image.Inline != null) {
                    var extents = _Image.Inline.Extent;
                    var cX = extents.Cx;
                    return cX / EnglishMetricUnitsPerInch * PixelsPerInch;
                } else if (_Image.Anchor != null) {
                    var extents = _Image.Anchor.Extent;
                    var cX = extents.Cx;
                    return cX / EnglishMetricUnitsPerInch * PixelsPerInch;
                }
                return null;
            }
            set {
                double emuWidth = value.Value * EnglishMetricUnitsPerInch / PixelsPerInch;
                if (_Image.Inline != null) {
                    var extents = _Image.Inline.Extent;
                    extents.Cx = (long)emuWidth;
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    picture.ShapeProperties.Transform2D.Extents.Cx = (Int64Value)emuWidth;
                } else if (_Image.Anchor != null) {
                    var extents = _Image.Anchor.Extent;
                    extents.Cx = (long)emuWidth;
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        picture.ShapeProperties.Transform2D.Extents.Cx = (Int64Value)emuWidth;
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
                if (_Image.Inline != null) {
                    var extents = _Image.Inline.Extent;
                    var cY = extents.Cy;
                    return cY / EnglishMetricUnitsPerInch * PixelsPerInch;
                } else if (_Image.Anchor != null) {
                    var extents = _Image.Anchor.Extent;
                    var cY = extents.Cy;
                    return cY / EnglishMetricUnitsPerInch * PixelsPerInch;
                }
                return null;
            }
            set {
                if (_Image.Inline != null) {
                    double emuHeight = value.Value * EnglishMetricUnitsPerInch / PixelsPerInch;
                    _Image.Inline.Extent.Cy = (Int64Value)emuHeight;
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    picture.ShapeProperties.Transform2D.Extents.Cy = (Int64Value)emuHeight;
                } else if (_Image.Anchor != null) {
                    double emuHeight = value.Value * EnglishMetricUnitsPerInch / PixelsPerInch;
                    _Image.Anchor.Extent.Cy = (Int64Value)emuHeight;
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        picture.ShapeProperties.Transform2D.Extents.Cy = (Int64Value)emuHeight;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the image width in EMUs.
        /// </summary>
        public double? EmuWidth {
            get {
                if (_Image.Inline != null) {
                    var extents = _Image.Inline.Extent;
                    return extents.Cx;
                } else if (_Image.Anchor != null) {
                    var extents = _Image.Anchor.Extent;
                    return extents.Cx;
                }
                return null;
            }
        }
        /// <summary>
        /// Gets the image height in EMUs.
        /// </summary>
        public double? EmuHeight {
            get {
                if (_Image.Inline != null) {
                    var extents = _Image.Inline.Extent;
                    return extents.Cy;
                } else if (_Image.Anchor != null) {
                    var extents = _Image.Anchor.Extent;
                    return extents.Cy;
                }
                return null;
            }
        }

        /// <summary>
        /// Gets or sets the shape type used to display the image.
        /// </summary>
        public ShapeTypeValues? Shape {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var presetGeometry = picture.ShapeProperties.GetFirstChild<PresetGeometry>();
                    return presetGeometry.Preset;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        var presetGeometry = picture.ShapeProperties.GetFirstChild<PresetGeometry>();
                        return presetGeometry.Preset;
                    }
                }

                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    var presetGeometry = picture.ShapeProperties.GetFirstChild<PresetGeometry>();
                    presetGeometry.Preset = value;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        var presetGeometry = picture.ShapeProperties.GetFirstChild<PresetGeometry>();
                        presetGeometry.Preset = value;
                    }
                }

            }
        }

        /// <summary>
        /// Microsoft Office does not seem to fully support this attribute, and ignores this setting.
        /// More information: http://officeopenxml.com/drwSp-SpPr.php
        /// </summary>
        public BlackWhiteModeValues? BlackWiteMode {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    return picture.ShapeProperties.BlackWhiteMode.Value;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        return picture.ShapeProperties.BlackWhiteMode.Value;
                    }
                }

                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();

                    if (value == null) {
                        // delete?
                    } else {
                        if (picture.ShapeProperties.BlackWhiteMode == null) {
                            picture.ShapeProperties.BlackWhiteMode = new EnumValue<BlackWhiteModeValues>();
                        }
                        picture.ShapeProperties.BlackWhiteMode.Value = value.Value;
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (value == null) {
                        // delete?
                    } else {
                        if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                            var picture = anchorGraphic.GraphicData
                                .GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                            if (picture.ShapeProperties.BlackWhiteMode == null) {
                                picture.ShapeProperties.BlackWhiteMode = new EnumValue<BlackWhiteModeValues>();
                            }

                            picture.ShapeProperties.BlackWhiteMode.Value = value.Value;
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
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    if (picture.ShapeProperties.Transform2D != null) {
                        if (picture.ShapeProperties.Transform2D.VerticalFlip != null) {
                            return picture.ShapeProperties.Transform2D.VerticalFlip.Value;
                        }
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        if (picture.ShapeProperties.Transform2D != null) {
                            if (picture.ShapeProperties.Transform2D.VerticalFlip != null) {
                                return picture.ShapeProperties.Transform2D.VerticalFlip.Value;
                            }
                        }
                    }
                }

                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    picture.ShapeProperties.Transform2D.VerticalFlip = value.Value;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        picture.ShapeProperties.Transform2D.VerticalFlip = value.Value;
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
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    if (picture.ShapeProperties.Transform2D != null) {
                        if (picture.ShapeProperties.Transform2D.HorizontalFlip != null) {
                            return picture.ShapeProperties.Transform2D.HorizontalFlip.Value;
                        }
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        if (picture.ShapeProperties.Transform2D != null) {
                            if (picture.ShapeProperties.Transform2D.HorizontalFlip != null) {
                                return picture.ShapeProperties.Transform2D.HorizontalFlip.Value;
                            }
                        }
                    }
                }

                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    picture.ShapeProperties.Transform2D.HorizontalFlip = value.Value;
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        picture.ShapeProperties.Transform2D.HorizontalFlip = value.Value;
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
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    if (picture.ShapeProperties.Transform2D.Rotation != null) {
                        return picture.ShapeProperties.Transform2D.Rotation / 10000;
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        if (picture.ShapeProperties.Transform2D.Rotation != null) {
                            return picture.ShapeProperties.Transform2D.Rotation / 10000;
                        }
                    }
                }

                return null;
            }
            set {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    if (value == null) {
                        picture.ShapeProperties.Transform2D.Rotation = null;
                    } else {
                        picture.ShapeProperties.Transform2D.Rotation = value.Value * 10000;
                    }
                } else if (_Image.Anchor != null) {
                    var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                    if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                        var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                        if (value == null) {
                            picture.ShapeProperties.Transform2D.Rotation = null;
                        } else {
                            picture.ShapeProperties.Transform2D.Rotation = value.Value * 10000;
                        }
                    }
                }
            }
        }

        private DocumentFormat.OpenXml.Drawing.Pictures.Picture? GetPicture() {
            if (_Image.Inline != null) {
                return _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
            }

            if (_Image.Anchor != null) {
                var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                    return anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                }
            }

            return null;
        }

        /// <summary>
        /// Gets or sets the number of EMUs to crop from the top of the image.
        /// </summary>
        public int? CropTop {
            get {
                var picture = GetPicture();
                return picture?.BlipFill?.SourceRectangle?.Top;
            }
            set {
                _cropTop = value;
                var picture = GetPicture();
                if (picture == null) return;

                if (picture.BlipFill.SourceRectangle == null && value != null) {
                    picture.BlipFill.SourceRectangle = new SourceRectangle();
                }

                if (picture.BlipFill.SourceRectangle != null) {
                    picture.BlipFill.SourceRectangle.Top = value;
                    if (value == null &&
                        picture.BlipFill.SourceRectangle.Left == null &&
                        picture.BlipFill.SourceRectangle.Right == null &&
                        picture.BlipFill.SourceRectangle.Bottom == null) {
                        picture.BlipFill.SourceRectangle.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the number of EMUs to crop from the bottom of the image.
        /// </summary>
        public int? CropBottom {
            get {
                var picture = GetPicture();
                return picture?.BlipFill?.SourceRectangle?.Bottom;
            }
            set {
                _cropBottom = value;
                var picture = GetPicture();
                if (picture == null) return;

                if (picture.BlipFill.SourceRectangle == null && value != null) {
                    picture.BlipFill.SourceRectangle = new SourceRectangle();
                }

                if (picture.BlipFill.SourceRectangle != null) {
                    picture.BlipFill.SourceRectangle.Bottom = value;
                    if (value == null &&
                        picture.BlipFill.SourceRectangle.Left == null &&
                        picture.BlipFill.SourceRectangle.Right == null &&
                        picture.BlipFill.SourceRectangle.Top == null) {
                        picture.BlipFill.SourceRectangle.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the number of EMUs to crop from the left side of the image.
        /// </summary>
        public int? CropLeft {
            get {
                var picture = GetPicture();
                return picture?.BlipFill?.SourceRectangle?.Left;
            }
            set {
                _cropLeft = value;
                var picture = GetPicture();
                if (picture == null) return;

                if (picture.BlipFill.SourceRectangle == null && value != null) {
                    picture.BlipFill.SourceRectangle = new SourceRectangle();
                }

                if (picture.BlipFill.SourceRectangle != null) {
                    picture.BlipFill.SourceRectangle.Left = value;
                    if (value == null &&
                        picture.BlipFill.SourceRectangle.Top == null &&
                        picture.BlipFill.SourceRectangle.Right == null &&
                        picture.BlipFill.SourceRectangle.Bottom == null) {
                        picture.BlipFill.SourceRectangle.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the number of EMUs to crop from the right side of the image.
        /// </summary>
        public int? CropRight {
            get {
                var picture = GetPicture();
                return picture?.BlipFill?.SourceRectangle?.Right;
            }
            set {
                _cropRight = value;
                var picture = GetPicture();
                if (picture == null) return;

                if (picture.BlipFill.SourceRectangle == null && value != null) {
                    picture.BlipFill.SourceRectangle = new SourceRectangle();
                }

                if (picture.BlipFill.SourceRectangle != null) {
                    picture.BlipFill.SourceRectangle.Right = value;
                    if (value == null &&
                        picture.BlipFill.SourceRectangle.Top == null &&
                        picture.BlipFill.SourceRectangle.Left == null &&
                        picture.BlipFill.SourceRectangle.Bottom == null) {
                        picture.BlipFill.SourceRectangle.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the number of centimeters to crop from the top of the image.
        /// </summary>
        public double? CropTopCentimeters {
            get {
                if (CropTop != null) {
                    return Helpers.ConvertEmusToCentimeters(CropTop.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    CropTop = Helpers.ConvertCentimetersToEmus(value.Value);
                } else {
                    CropTop = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the number of centimeters to crop from the bottom of the image.
        /// </summary>
        public double? CropBottomCentimeters {
            get {
                if (CropBottom != null) {
                    return Helpers.ConvertEmusToCentimeters(CropBottom.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    CropBottom = Helpers.ConvertCentimetersToEmus(value.Value);
                } else {
                    CropBottom = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the number of centimeters to crop from the left side of the image.
        /// </summary>
        public double? CropLeftCentimeters {
            get {
                if (CropLeft != null) {
                    return Helpers.ConvertEmusToCentimeters(CropLeft.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    CropLeft = Helpers.ConvertCentimetersToEmus(value.Value);
                } else {
                    CropLeft = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the number of centimeters to crop from the right side of the image.
        /// </summary>
        public double? CropRightCentimeters {
            get {
                if (CropRight != null) {
                    return Helpers.ConvertEmusToCentimeters(CropRight.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    CropRight = Helpers.ConvertCentimetersToEmus(value.Value);
                } else {
                    CropRight = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets the image transparency percentage (0-100).
        /// </summary>
        public int? Transparency {
            get {
                var blip = GetBlip();
                if (blip != null) {
                    var alpha = blip.GetFirstChild<AlphaModulationFixed>();
                    if (alpha != null) {
                        return (int)((100000 - alpha.Amount.Value) / 1000);
                    }
                }
                return null;
            }
            set {
                if (value is < 0 or > 100)
                    throw new ArgumentOutOfRangeException(nameof(value), "Transparency must be between 0 and 100.");

                var blip = GetBlip();
                if (blip == null) return;

                var alpha = blip.GetFirstChild<AlphaModulationFixed>();
                if (value == null) {
                    alpha?.Remove();
                    return;
                }

                if (alpha == null) {
                    alpha = new AlphaModulationFixed();
                    blip.Append(alpha);
                }
                alpha.Amount = 100000 - value.Value * 1000;
            }
        }

        /// <summary>
        /// Gets or sets the image's wrap text.
        /// </summary>
        public WrapTextImage? WrapText {
            get => WordWrapTextImage.GetWrapTextImage(_Image.Anchor, _Image.Inline);
            set => WordWrapTextImage.SetWrapTextImage(_Image, _Image.Anchor, _Image.Inline, value);
        }

        /// <summary>
        /// Gets or sets how the image should fill its bounding box. Default is Stretch.
        /// </summary>
        public ImageFillMode FillMode {
            get {
                var picture = GetPicture();
                var blipFill = picture?.BlipFill;
                if (blipFill != null && blipFill.GetFirstChild<Tile>() != null) {
                    _fillMode = ImageFillMode.Tile;
                } else {
                    _fillMode = ImageFillMode.Stretch;
                }
                return _fillMode;
            }
            set {
                _fillMode = value;
                var picture = GetPicture();
                if (picture == null) return;

                var blipFill = picture.BlipFill;
                var tile = blipFill.GetFirstChild<Tile>();
                var stretch = blipFill.GetFirstChild<Stretch>();

                if (value == ImageFillMode.Stretch) {
                    tile?.Remove();
                    if (stretch == null) {
                        blipFill.AppendChild(new Stretch(new FillRectangle()));
                    }
                } else {
                    stretch?.Remove();
                    if (tile == null) {
                        blipFill.AppendChild(new Tile());
                    }
                }
            }
        }

        /// <summary>
        /// Specifies whether to use the image's local DPI setting when rendering.
        /// </summary>
        public bool? UseLocalDpi {
            get {
                var blip = GetBlip();
                var extList = blip?.GetFirstChild<BlipExtensionList>();
                var useLocalDpi = extList?
                    .OfType<BlipExtension>()
                    .FirstOrDefault(e => e.Uri == "{28A0092B-C50C-407E-A947-70E740481C1C}")?
                    .GetFirstChild<DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi>();
                _useLocalDpi = useLocalDpi?.Val;
                return _useLocalDpi;
            }
            set {
                _useLocalDpi = value;
                var blip = GetBlip();
                if (blip == null) return;

                var extList = blip.GetFirstChild<BlipExtensionList>();
                if (extList == null && value != null) {
                    extList = new BlipExtensionList();
                    blip.Append(extList);
                }

                if (extList != null) {
                    var extension = extList
                        .OfType<BlipExtension>()
                        .FirstOrDefault(e => e.Uri == "{28A0092B-C50C-407E-A947-70E740481C1C}");
                    if (extension == null && value != null) {
                        extension = new BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
                        extList.Append(extension);
                    }
                    if (extension != null) {
                        var useLocalDpiElement = extension.GetFirstChild<DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi>();
                        if (value == null) {
                            useLocalDpiElement?.Remove();
                        } else {
                            if (useLocalDpiElement == null) {
                                useLocalDpiElement = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi();
                                useLocalDpiElement.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                                extension.Append(useLocalDpiElement);
                            }
                            useLocalDpiElement.Val = value.Value;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets the anchor location of the image.
        /// </summary>
        public (int X, int Y) Location {
            get {
                if (_Image.Anchor != null) {
                    var hPos = _Image.Anchor.HorizontalPosition;
                    var vPos = _Image.Anchor.VerticalPosition;
                    int x = 0;
                    int y = 0;
                    if (hPos?.PositionOffset != null) {
                        int.TryParse(hPos.PositionOffset.Text, out x);
                    }
                    if (vPos?.PositionOffset != null) {
                        int.TryParse(vPos.PositionOffset.Text, out y);
                    }
                    return (x, y);
                }

                return (0, 0);
            }
        }

        /// <summary>
        /// Indicates whether resizing should be relative to the original size.
        /// </summary>
        public bool? PreferRelativeResize {
            get {
                var pic = GetPicture();
                var nv = pic?.NonVisualPictureProperties?.NonVisualPictureDrawingProperties;
                _preferRelativeResize = nv?.PreferRelativeResize;
                return _preferRelativeResize;
            }
            set {
                _preferRelativeResize = value;
                var pic = GetPicture();
                if (pic == null) return;
                var nv = pic.NonVisualPictureProperties.NonVisualPictureDrawingProperties;
                if (nv == null && value != null) {
                    nv = new Pic.NonVisualPictureDrawingProperties();
                    pic.NonVisualPictureProperties.Append(nv);
                }
                if (nv != null) {
                    nv.PreferRelativeResize = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets whether the aspect ratio is locked.
        /// </summary>
        public bool? NoChangeAspect {
            get {
                var locks = GetPicture()?.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.PictureLocks;
                _noChangeAspect = locks?.NoChangeAspect;
                return _noChangeAspect;
            }
            set {
                _noChangeAspect = value;
                SetPictureLock(l => l.NoChangeAspect = value);
            }
        }

        /// <summary>
        /// Gets or sets whether the image cannot be cropped.
        /// </summary>
        public bool? NoCrop {
            get {
                var locks = GetPicture()?.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.PictureLocks;
                _noCrop = locks?.NoCrop;
                return _noCrop;
            }
            set {
                _noCrop = value;
                SetPictureLock(l => l.NoCrop = value);
            }
        }

        /// <summary>
        /// Gets or sets whether the image is locked from moving.
        /// </summary>
        public bool? NoMove {
            get {
                var locks = GetPicture()?.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.PictureLocks;
                _noMove = locks?.NoMove;
                return _noMove;
            }
            set {
                _noMove = value;
                SetPictureLock(l => l.NoMove = value);
            }
        }

        /// <summary>
        /// Gets or sets whether the image cannot be resized.
        /// </summary>
        public bool? NoResize {
            get {
                var locks = GetPicture()?.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.PictureLocks;
                _noResize = locks?.NoResize;
                return _noResize;
            }
            set {
                _noResize = value;
                SetPictureLock(l => l.NoResize = value);
            }
        }

        /// <summary>
        /// Gets or sets whether the image cannot be rotated.
        /// </summary>
        public bool? NoRot {
            get {
                var locks = GetPicture()?.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.PictureLocks;
                _noRotation = locks?.NoRotation;
                return _noRotation;
            }
            set {
                _noRotation = value;
                SetPictureLock(l => l.NoRotation = value);
            }
        }

        /// <summary>
        /// Gets or sets whether the image cannot be selected.
        /// </summary>
        public bool? NoSelect {
            get {
                var locks = GetPicture()?.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.PictureLocks;
                _noSelection = locks?.NoSelection;
                return _noSelection;
            }
            set {
                _noSelection = value;
                SetPictureLock(l => l.NoSelection = value);
            }
        }

        /// <summary>
        /// Sets picture lock property using provided action.
        /// </summary>
        private void SetPictureLock(Action<A.PictureLocks> setter) {
            var pic = GetPicture();
            if (pic == null) return;
            var nv = pic.NonVisualPictureProperties.NonVisualPictureDrawingProperties;
            if (nv == null) {
                nv = new Pic.NonVisualPictureDrawingProperties();
                pic.NonVisualPictureProperties.Append(nv);
            }
            var locks = nv.PictureLocks;
            if (locks == null) {
                locks = new A.PictureLocks();
                nv.Append(locks);
            }
            setter(locks);
        }

        /// <summary>
        /// Sets a fixed opacity value for the image.
        /// </summary>
        public int? FixedOpacity {
            get {
                var blip = GetBlip();
                var ar = blip?.GetFirstChild<AlphaReplace>();
                _fixedOpacity = ar != null ? (int?)(ar.Alpha.Value / 1000) : null;
                return _fixedOpacity;
            }
            set {
                if (value is < 0 or > 100)
                    throw new ArgumentOutOfRangeException(nameof(value), "Opacity must be between 0 and 100.");

                _fixedOpacity = value;
                var blip = GetBlip();
                if (blip == null) return;
                var ar = blip.GetFirstChild<AlphaReplace>();
                if (value == null) {
                    ar?.Remove();
                    return;
                }
                if (ar == null) {
                    ar = new AlphaReplace();
                    blip.Append(ar);
                }
                ar.Alpha = value.Value * 1000;
            }
        }

        /// <summary>
        /// Gets or sets the alpha inversion color in hex.
        /// </summary>
        public string AlphaInversionColorHex {
            get {
                var blip = GetBlip();
                var ai = blip?.GetFirstChild<AlphaInverse>();
                _alphaInversionColorHex = ai?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _alphaInversionColorHex;
            }
            set {
                _alphaInversionColorHex = value;
                var blip = GetBlip();
                if (blip == null) return;
                var ai = blip.GetFirstChild<AlphaInverse>();
                if (value == null) {
                    ai?.Remove();
                    return;
                }
                if (ai == null) {
                    ai = new AlphaInverse();
                    blip.Append(ai);
                }
                var clr = ai.GetFirstChild<RgbColorModelHex>();
                if (clr == null) {
                    ai.RemoveAllChildren();
                    clr = new RgbColorModelHex();
                    ai.Append(clr);
                }
                clr.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the alpha inversion color.
        /// </summary>
        public SixLabors.ImageSharp.Color? AlphaInversionColor {
            get {
                if (AlphaInversionColorHex == null) return (SixLabors.ImageSharp.Color?)null;
                return Helpers.ParseColor(AlphaInversionColorHex);
            }
            set { AlphaInversionColorHex = value?.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the threshold for black and white effect.
        /// </summary>
        public int? BlackWhiteThreshold {
            get {
                var blip = GetBlip();
                var bi = blip?.GetFirstChild<BiLevel>();
                _blackWhiteThreshold = bi != null ? (int?)(bi.Threshold.Value / 1000) : null;
                return _blackWhiteThreshold;
            }
            set {
                if (value is < 0 or > 100)
                    throw new ArgumentOutOfRangeException(nameof(value), "Threshold must be between 0 and 100.");
                _blackWhiteThreshold = value;
                var blip = GetBlip();
                if (blip == null) return;
                var bi = blip.GetFirstChild<BiLevel>();
                if (value == null) {
                    bi?.Remove();
                    return;
                }
                if (bi == null) {
                    bi = new BiLevel();
                    blip.Append(bi);
                }
                bi.Threshold = value.Value * 1000;
            }
        }

        /// <summary>
        /// Gets or sets the blur radius in EMUs.
        /// </summary>
        public int? BlurRadius {
            get {
                var blip = GetBlip();
                var blur = blip?.GetFirstChild<Blur>();
                _blurRadius = blur?.Radius != null ? (int?)blur.Radius.Value : null;
                return _blurRadius;
            }
            set {
                _blurRadius = value;
                var blip = GetBlip();
                if (blip == null) return;
                var blur = blip.GetFirstChild<Blur>();
                if (value == null && _blurGrow == null) {
                    blur?.Remove();
                    return;
                }
                if (blur == null) {
                    blur = new Blur();
                    blip.Append(blur);
                }
                blur.Radius = value != null ? new Int64Value((long)value.Value) : null;
                blur.Grow = _blurGrow ?? false;
            }
        }

        /// <summary>
        /// Gets or sets whether the blur grows the image's bounds.
        /// </summary>
        public bool? BlurGrow {
            get {
                var blip = GetBlip();
                var blur = blip?.GetFirstChild<Blur>();
                _blurGrow = blur?.Grow;
                return _blurGrow;
            }
            set {
                _blurGrow = value;
                var blip = GetBlip();
                if (blip == null) return;
                var blur = blip.GetFirstChild<Blur>();
                if (value == null && _blurRadius == null) {
                    blur?.Remove();
                    return;
                }
                if (blur == null) {
                    blur = new Blur();
                    blip.Append(blur);
                }
                blur.Radius = _blurRadius != null ? new Int64Value((long)_blurRadius.Value) : null;
                blur.Grow = value ?? false;
            }
        }

        /// <summary>
        /// Gets or sets the source color to change from in hex.
        /// </summary>
        public string ColorChangeFromHex {
            get {
                var blip = GetBlip();
                var cc = blip?.GetFirstChild<ColorChange>();
                _colorChangeFromHex = cc?.ColorFrom?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _colorChangeFromHex;
            }
            set {
                _colorChangeFromHex = value;
                UpdateColorChange();
            }
        }

        /// <summary>
        /// Gets or sets the target color to change to in hex.
        /// </summary>
        public string ColorChangeToHex {
            get {
                var blip = GetBlip();
                var cc = blip?.GetFirstChild<ColorChange>();
                _colorChangeToHex = cc?.ColorTo?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _colorChangeToHex;
            }
            set {
                _colorChangeToHex = value;
                UpdateColorChange();
            }
        }

        /// <summary>
        /// Gets or sets the source color to change from.
        /// </summary>
        public SixLabors.ImageSharp.Color? ColorChangeFrom {
            get {
                return ColorChangeFromHex == null ? (SixLabors.ImageSharp.Color?)null : Helpers.ParseColor(ColorChangeFromHex);
            }
            set { ColorChangeFromHex = value?.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the target color to change to.
        /// </summary>
        public SixLabors.ImageSharp.Color? ColorChangeTo {
            get {
                return ColorChangeToHex == null ? (SixLabors.ImageSharp.Color?)null : Helpers.ParseColor(ColorChangeToHex);
            }
            set { ColorChangeToHex = value?.ToHexColor(); }
        }

        private void UpdateColorChange() {
            var blip = GetBlip();
            if (blip == null) return;
            var cc = blip.GetFirstChild<ColorChange>();
            if (_colorChangeFromHex == null && _colorChangeToHex == null) {
                cc?.Remove();
                return;
            }
            if (cc == null) {
                cc = new ColorChange();
                blip.Append(cc);
            }
            if (_colorChangeFromHex != null) {
                var from = cc.ColorFrom ?? new ColorFrom();
                var clr = from.GetFirstChild<RgbColorModelHex>() ?? new RgbColorModelHex();
                clr.Val = _colorChangeFromHex;
                if (from.FirstChild == null) from.Append(clr);
                cc.ColorFrom = from;
            }
            if (_colorChangeToHex != null) {
                var to = cc.ColorTo ?? new ColorTo();
                var clr = to.GetFirstChild<RgbColorModelHex>() ?? new RgbColorModelHex();
                clr.Val = _colorChangeToHex;
                if (to.FirstChild == null) to.Append(clr);
                cc.ColorTo = to;
            }
        }

        /// <summary>
        /// Gets or sets a color replacement hex value.
        /// </summary>
        public string ColorReplacementHex {
            get {
                var blip = GetBlip();
                var cr = blip?.GetFirstChild<ColorReplacement>();
                _colorReplacementHex = cr?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _colorReplacementHex;
            }
            set {
                _colorReplacementHex = value;
                var blip = GetBlip();
                if (blip == null) return;
                var cr = blip.GetFirstChild<ColorReplacement>();
                if (value == null) {
                    cr?.Remove();
                    return;
                }
                if (cr == null) {
                    cr = new ColorReplacement();
                    blip.Append(cr);
                }
                var clr = cr.GetFirstChild<RgbColorModelHex>();
                if (clr == null) {
                    cr.RemoveAllChildren();
                    clr = new RgbColorModelHex();
                    cr.Append(clr);
                }
                clr.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets a color replacement.
        /// </summary>
        public SixLabors.ImageSharp.Color? ColorReplacement {
            get {
                return ColorReplacementHex == null ? (SixLabors.ImageSharp.Color?)null : Helpers.ParseColor(ColorReplacementHex);
            }
            set { ColorReplacementHex = value?.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the first duotone color in hex.
        /// </summary>
        public string DuotoneColor1Hex {
            get {
                var blip = GetBlip();
                var duo = blip?.GetFirstChild<Duotone>();
                _duotoneColor1Hex = duo?.GetFirstChild<RgbColorModelHex>()?.Val;
                return _duotoneColor1Hex;
            }
            set {
                _duotoneColor1Hex = value;
                UpdateDuotone();
            }
        }

        /// <summary>
        /// Gets or sets the second duotone color in hex.
        /// </summary>
        public string DuotoneColor2Hex {
            get {
                var blip = GetBlip();
                var duo = blip?.GetFirstChild<Duotone>();
                _duotoneColor2Hex = duo?.Elements<RgbColorModelHex>().Skip(1).FirstOrDefault()?.Val;
                return _duotoneColor2Hex;
            }
            set {
                _duotoneColor2Hex = value;
                UpdateDuotone();
            }
        }

        /// <summary>
        /// Gets or sets the first duotone color.
        /// </summary>
        public SixLabors.ImageSharp.Color? DuotoneColor1 {
            get {
                return DuotoneColor1Hex == null ? (SixLabors.ImageSharp.Color?)null : Helpers.ParseColor(DuotoneColor1Hex);
            }
            set { DuotoneColor1Hex = value?.ToHexColor(); }
        }

        /// <summary>
        /// Gets or sets the second duotone color.
        /// </summary>
        public SixLabors.ImageSharp.Color? DuotoneColor2 {
            get {
                return DuotoneColor2Hex == null ? (SixLabors.ImageSharp.Color?)null : Helpers.ParseColor(DuotoneColor2Hex);
            }
            set { DuotoneColor2Hex = value?.ToHexColor(); }
        }

        private void UpdateDuotone() {
            var blip = GetBlip();
            if (blip == null) return;
            var duo = blip.GetFirstChild<Duotone>();
            if (_duotoneColor1Hex == null && _duotoneColor2Hex == null) {
                duo?.Remove();
                return;
            }
            if (duo == null) {
                duo = new Duotone();
                blip.Append(duo);
            }
            duo.RemoveAllChildren();
            if (_duotoneColor1Hex != null)
                duo.Append(new RgbColorModelHex { Val = _duotoneColor1Hex });
            if (_duotoneColor2Hex != null)
                duo.Append(new RgbColorModelHex { Val = _duotoneColor2Hex });
        }

        /// <summary>
        /// Gets or sets whether the image is displayed in grayscale.
        /// </summary>
        public bool? GrayScale {
            get {
                var blip = GetBlip();
                var gs = blip?.GetFirstChild<Grayscale>();
                _grayScale = gs != null;
                return _grayScale;
            }
            set {
                _grayScale = value;
                var blip = GetBlip();
                if (blip == null) return;
                var gs = blip.GetFirstChild<Grayscale>();
                if (value == true && gs == null) {
                    blip.Append(new Grayscale());
                } else if (value != true) {
                    gs?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the brightness adjustment.
        /// </summary>
        public int? LuminanceBrightness {
            get {
                var blip = GetBlip();
                var lum = blip?.GetFirstChild<LuminanceEffect>();
                _luminanceBrightness = lum != null ? lum.Brightness : null;
                if (_luminanceBrightness != null) _luminanceBrightness /= 1000;
                return _luminanceBrightness;
            }
            set {
                _luminanceBrightness = value;
                UpdateLuminance();
            }
        }

        /// <summary>
        /// Gets or sets the contrast adjustment.
        /// </summary>
        public int? LuminanceContrast {
            get {
                var blip = GetBlip();
                var lum = blip?.GetFirstChild<LuminanceEffect>();
                _luminanceContrast = lum != null ? lum.Contrast : null;
                if (_luminanceContrast != null) _luminanceContrast /= 1000;
                return _luminanceContrast;
            }
            set {
                _luminanceContrast = value;
                UpdateLuminance();
            }
        }

        private void UpdateLuminance() {
            var blip = GetBlip();
            if (blip == null) return;
            var lum = blip.GetFirstChild<LuminanceEffect>();
            if (_luminanceBrightness == null && _luminanceContrast == null) {
                lum?.Remove();
                return;
            }
            if (lum == null) {
                lum = new LuminanceEffect();
                blip.Append(lum);
            }
            lum.Brightness = _luminanceBrightness != null ? new Int32Value(_luminanceBrightness.Value * 1000) : null;
            lum.Contrast = _luminanceContrast != null ? new Int32Value(_luminanceContrast.Value * 1000) : null;
        }

        /// <summary>
        /// Gets or sets the tint amount.
        /// </summary>
        public int? TintAmount {
            get {
                var blip = GetBlip();
                var tint = blip?.GetFirstChild<TintEffect>();
                _tintAmount = tint?.Amount != null ? tint.Amount / 1000 : null;
                return _tintAmount;
            }
            set {
                _tintAmount = value;
                UpdateTint();
            }
        }

        /// <summary>
        /// Gets or sets the tint hue.
        /// </summary>
        public int? TintHue {
            get {
                var blip = GetBlip();
                var tint = blip?.GetFirstChild<TintEffect>();
                _tintHue = tint?.Hue != null ? tint.Hue / 60000 : null;
                return _tintHue;
            }
            set {
                _tintHue = value;
                UpdateTint();
            }
        }

        private void UpdateTint() {
            var blip = GetBlip();
            if (blip == null) return;
            var tint = blip.GetFirstChild<TintEffect>();
            if (_tintAmount == null && _tintHue == null) {
                tint?.Remove();
                return;
            }
            if (tint == null) {
                tint = new TintEffect();
                blip.Append(tint);
            }
            tint.Amount = _tintAmount != null ? new Int32Value(_tintAmount.Value * 1000) : null;
            tint.Hue = _tintHue != null ? new Int32Value(_tintHue.Value * 60000) : null;
        }

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

            var nvDrawingProps = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
            {
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

            blipFlip.Append(blip);
            if (_cropTop != null || _cropBottom != null || _cropLeft != null || _cropRight != null) {
                var srcRect = new SourceRectangle();
                if (_cropTop != null) srcRect.Top = _cropTop;
                if (_cropBottom != null) srcRect.Bottom = _cropBottom;
                if (_cropLeft != null) srcRect.Left = _cropLeft;
                if (_cropRight != null) srcRect.Right = _cropRight;
                blipFlip.Append(srcRect);
            }

            if (_fillMode == ImageFillMode.Stretch) {
                blipFlip.Append(new Stretch(new FillRectangle()));
            } else {
                if (blipFlip.GetFirstChild<Tile>() == null) {
                    blipFlip.AppendChild(new Tile());
                }
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
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                return picture?.BlipFill?.Blip;
            } else if (_Image.Anchor != null) {
                var anchorGraphic = _Image.Anchor.OfType<Graphic>().FirstOrDefault();
                if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                    var picture = anchorGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    return picture?.BlipFill?.Blip;
                }
            }
            return null;
        }

        /// <summary>
        /// Wraps an existing drawing as a WordImage.
        /// </summary>
        public WordImage(WordDocument document, Drawing drawing) {
            _document = document;
            _Image = drawing;

            var initialBlip = GetBlip();
            if (initialBlip != null) {
                if (initialBlip.Link != null) {
                    _externalRelationshipId = initialBlip.Link;
                } else if (initialBlip.Embed != null) {
                    var part = GetContainingPart();
                    _imagePart = part.GetPartById(initialBlip.Embed) as ImagePart;
                }
            }

            var picture = GetPicture();
            if (picture != null) {
                var nv = picture.NonVisualPictureProperties;
                if (nv != null) {
                    _title = nv.NonVisualDrawingProperties.Title;
                    _hidden = nv.NonVisualDrawingProperties.Hidden?.Value;
                    var nvPic = nv.NonVisualPictureDrawingProperties;
                    if (nvPic != null) {
                        _preferRelativeResize = nvPic.PreferRelativeResize?.Value;
                        var locks = nvPic.PictureLocks;
                        if (locks != null) {
                            _noChangeAspect = locks.NoChangeAspect?.Value;
                            _noCrop = locks.NoCrop?.Value;
                            _noMove = locks.NoMove?.Value;
                            _noResize = locks.NoResize?.Value;
                            _noRotation = locks.NoRotation?.Value;
                            _noSelection = locks.NoSelection?.Value;
                        }
                    }
                }

                var pictureBlip = picture.BlipFill?.Blip;
                if (pictureBlip != null) {
                    var ar = pictureBlip.GetFirstChild<AlphaReplace>();
                    if (ar != null) _fixedOpacity = (int?)(ar.Alpha.Value / 1000);
                    var ai = pictureBlip.GetFirstChild<AlphaInverse>();
                    _alphaInversionColorHex = ai?.GetFirstChild<RgbColorModelHex>()?.Val;
                    var bi = pictureBlip.GetFirstChild<BiLevel>();
                    _blackWhiteThreshold = bi != null ? (int?)(bi.Threshold.Value / 1000) : null;
                    var blur = pictureBlip.GetFirstChild<Blur>();
                    if (blur != null) { _blurRadius = blur.Radius != null ? (int?)blur.Radius.Value : null; _blurGrow = blur.Grow; }
                    var cc = pictureBlip.GetFirstChild<ColorChange>();
                    if (cc != null) {
                        _colorChangeFromHex = cc.ColorFrom?.GetFirstChild<RgbColorModelHex>()?.Val;
                        _colorChangeToHex = cc.ColorTo?.GetFirstChild<RgbColorModelHex>()?.Val;
                    }
                    var cr = pictureBlip.GetFirstChild<ColorReplacement>();
                    _colorReplacementHex = cr?.GetFirstChild<RgbColorModelHex>()?.Val;
                    var duo = pictureBlip.GetFirstChild<Duotone>();
                    if (duo != null) {
                        _duotoneColor1Hex = duo.GetFirstChild<RgbColorModelHex>()?.Val;
                        _duotoneColor2Hex = duo.Elements<RgbColorModelHex>().Skip(1).FirstOrDefault()?.Val;
                    }
                    _grayScale = pictureBlip.GetFirstChild<Grayscale>() != null;
                    var lum = pictureBlip.GetFirstChild<LuminanceEffect>();
                    if (lum != null) {
                        _luminanceBrightness = lum.Brightness != null ? (int?)(lum.Brightness.Value / 1000) : null;
                        _luminanceContrast = lum.Contrast != null ? (int?)(lum.Contrast.Value / 1000) : null;
                    }
                    var tint = pictureBlip.GetFirstChild<TintEffect>();
                    if (tint != null) {
                        _tintAmount = tint.Amount != null ? (int?)(tint.Amount.Value / 1000) : null;
                        _tintHue = tint.Hue != null ? (int?)(tint.Hue.Value / 60000) : null;
                    }
                    var ext = pictureBlip.GetFirstChild<BlipExtensionList>()?.OfType<BlipExtension>()
                        .FirstOrDefault(e => e.Uri == "{28A0092B-C50C-407E-A947-70E740481C1C}");
                    _useLocalDpi = ext?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi>()?.Val?.Value;
                }
            }
        }

        /// <summary>
        /// Extract image from Word Document and save it to file
        /// </summary>
        /// <param name="fileToSave"></param>
        public void SaveToFile(string fileToSave) {
            if (_imagePart == null) {
                throw new InvalidOperationException("Image is linked externally and cannot be saved.");
            }

            if (File.Exists(fileToSave) && new FileInfo(fileToSave).IsReadOnly) {
                throw new IOException($"Failed to save to '{fileToSave}'. The file is read-only.");
            }

            var directory = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(fileToSave));
            if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory)) {
                var dirInfo = new DirectoryInfo(directory);
                if (dirInfo.Attributes.HasFlag(FileAttributes.ReadOnly)) {
                    throw new IOException($"Failed to save to '{fileToSave}'. The directory is read-only.");
                }
            }

            try {
                using (FileStream outputFileStream = new FileStream(fileToSave, FileMode.Create)) {
                    var stream = _imagePart.GetStream();
                    stream.CopyTo(outputFileStream);
                    stream.Close();
                }
            } catch (UnauthorizedAccessException ex) {
                throw new IOException($"Failed to save to '{fileToSave}'. Access denied or path is read-only.", ex);
            }
        }

        /// <summary>
        /// Remove image from a Word Document
        /// </summary>
        public void Remove() {
            if (_imagePart != null) {
                OpenXmlElement parent = _Image.Parent;
                while (parent != null && parent is not Body && parent is not Header && parent is not Footer) {
                    parent = parent.Parent;
                }

                OpenXmlPart part = _document._wordprocessingDocument.MainDocumentPart;
                if (parent is Header header) {
                    part = header.HeaderPart;
                } else if (parent is Footer footer) {
                    part = footer.FooterPart;
                }

                part.DeletePart(_imagePart);
                _imagePart = null;
            } else if (!string.IsNullOrEmpty(_externalRelationshipId)) {
                OpenXmlElement parent = _Image.Parent;
                while (parent != null && parent is not Body && parent is not Header && parent is not Footer) {
                    parent = parent.Parent;
                }

                OpenXmlPart part = _document._wordprocessingDocument.MainDocumentPart;
                if (parent is Header header) {
                    part = header.HeaderPart;
                } else if (parent is Footer footer) {
                    part = footer.FooterPart;
                }

                var rel = part.ExternalRelationships.FirstOrDefault(r => r.Id == _externalRelationshipId);
                if (rel != null) {
                    part.DeleteExternalRelationship(rel);
                }
                _externalRelationshipId = null;
            }

            if (this._Image != null) {
                this._Image.Remove();
            }
        }

        private void AddImage(WordDocument document, WordParagraph paragraph, Stream imageStream, string fileName, double? width, double? height, ShapeTypeValues shape, BlipCompressionValues compressionQuality, string description, WrapTextImage wrapImage) {
            _document = document;
            var imageLocation = AddImageToLocation(document, paragraph, imageStream, fileName, width, height);

            this._imagePart = imageLocation.ImagePart;

            //calculate size in emu
            double emuWidth = imageLocation.Width * EnglishMetricUnitsPerInch / PixelsPerInch;
            double emuHeight = imageLocation.Height * EnglishMetricUnitsPerInch / PixelsPerInch;

            var drawing = new Drawing();

            if (wrapImage == WrapTextImage.InLineWithText) {
                var inline = GetInline(emuWidth, emuHeight, imageLocation.ImageName, fileName, imageLocation.RelationshipId, shape, compressionQuality, description);
                drawing.Append(inline);
            } else {
                var graphic = GetGraphic(emuWidth, emuHeight, fileName, imageLocation.RelationshipId, shape, compressionQuality, description);
                var anchor = GetAnchor(emuWidth, emuHeight, graphic, imageLocation.ImageName, description, wrapImage);
                drawing.Append(anchor);
            }
            this._Image = drawing;
        }

        internal static WordImageLocation AddImageToLocation(
            WordDocument document,
            WordParagraph paragraph,
            Stream imageStream,
            string fileName,
            double? width = null,
            double? height = null
        ) {

            // Size - https://stackoverflow.com/questions/8082980/inserting-image-into-docx-using-openxml-and-setting-the-size
            // if widht/height are not set we check ourselves
            // but probably will need better way
            var imageCharacteristics = Helpers.GetImageCharacteristics(imageStream, fileName);
            if (width == null || height == null) {
                if (imageCharacteristics.Width == 0 || imageCharacteristics.Height == 0) {
                    throw new ArgumentException("Width and height must be provided for this image type.");
                }
                width = imageCharacteristics.Width;
                height = imageCharacteristics.Height;
            }


            var imagePartType = imageCharacteristics.Type;
            var imageName = System.IO.Path.GetFileNameWithoutExtension(fileName);

            ImagePart imagePart;
            string relationshipId;
            var location = paragraph.Location();
            if (location.GetType() == typeof(Header)) {
                var part = ((Header)location).HeaderPart;
                imagePart = part.AddImagePart(imagePartType.ToOpenXmlImagePartType());
                relationshipId = part.GetIdOfPart(imagePart);
            } else if (location.GetType() == typeof(Footer)) {
                var part = ((Footer)location).FooterPart;
                imagePart = part.AddImagePart(imagePartType.ToOpenXmlImagePartType());
                relationshipId = part.GetIdOfPart(imagePart);
            } else if (location.GetType() == typeof(Document)) {
                var part = document._wordprocessingDocument.MainDocumentPart;
                imagePart = part.AddImagePart(imagePartType.ToOpenXmlImagePartType());
                relationshipId = part.GetIdOfPart(imagePart);
            } else {
                throw new InvalidOperationException("Paragraph is not in document or header or footer. This is weird. Probably a bug.");
            }

            imagePart.FeedData(imageStream);

            return new WordImageLocation() {
                ImagePart = imagePart,
                RelationshipId = relationshipId,
                Width = width.Value,
                Height = height.Value,
                ImageName = imageName
            };
        }

        private void AddExternalImage(WordDocument document, WordParagraph paragraph, Uri uri, double width, double height, ShapeTypeValues shape, BlipCompressionValues compressionQuality, string description, WrapTextImage wrapImage) {
            _document = document;

            var location = paragraph.Location();
            ExternalRelationship rel;
            if (location is Header header) {
                rel = header.HeaderPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", uri);
            } else if (location is Footer footer) {
                rel = footer.FooterPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", uri);
            } else {
                rel = document._wordprocessingDocument.MainDocumentPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", uri);
            }

            _externalRelationshipId = rel.Id;

            double emuWidth = width * EnglishMetricUnitsPerInch / PixelsPerInch;
            double emuHeight = height * EnglishMetricUnitsPerInch / PixelsPerInch;

            var drawing = new Drawing();
            if (wrapImage == WrapTextImage.InLineWithText) {
                var inline = GetInline(emuWidth, emuHeight, System.IO.Path.GetFileNameWithoutExtension(uri.ToString()), System.IO.Path.GetFileName(uri.ToString()), rel.Id, shape, compressionQuality, description, true);
                drawing.Append(inline);
            } else {
                var graphic = GetGraphic(emuWidth, emuHeight, System.IO.Path.GetFileName(uri.ToString()), rel.Id, shape, compressionQuality, description, true);
                var anchor = GetAnchor(emuWidth, emuHeight, graphic, System.IO.Path.GetFileNameWithoutExtension(uri.ToString()), description, wrapImage);
                drawing.Append(anchor);
            }
            _Image = drawing;
        }

        private OpenXmlPart GetContainingPart() {
            OpenXmlElement parent = _Image.Parent;
            while (parent != null && parent is not Body && parent is not Header && parent is not Footer) {
                parent = parent.Parent;
            }

            if (parent is Header header) {
                return header.HeaderPart;
            }

            if (parent is Footer footer) {
                return footer.FooterPart;
            }

            return _document._wordprocessingDocument.MainDocumentPart;
        }
    }
}
