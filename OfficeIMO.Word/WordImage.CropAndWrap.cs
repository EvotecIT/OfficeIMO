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
        /// Gets or sets the number of EMUs to crop from the top of the image.
        /// </summary>
        public int? CropTop {
            get {
                var picture = GetPicture();
                return picture?.BlipFill?.SourceRectangle?.Top?.Value;
            }
            set {
                _cropTop = value;
                var picture = GetPicture();
                if (picture == null) return;

                if (picture.BlipFill?.SourceRectangle == null && value != null) {
                    if (picture.BlipFill == null) {
                        picture.BlipFill = new Pic.BlipFill();
                    }
                    if (picture.BlipFill.SourceRectangle == null) {
                        picture.BlipFill.SourceRectangle = new A.SourceRectangle();
                    }
                }

                if (picture.BlipFill != null && picture.BlipFill.SourceRectangle != null) {
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
                return picture?.BlipFill?.SourceRectangle?.Bottom?.Value;
            }
            set {
                _cropBottom = value;
                var picture = GetPicture();
                if (picture == null) return;

                if (picture.BlipFill?.SourceRectangle == null && value != null) {
                    if (picture.BlipFill == null) {
                        picture.BlipFill = new Pic.BlipFill();
                    }
                    if (picture.BlipFill.SourceRectangle == null) {
                        picture.BlipFill.SourceRectangle = new A.SourceRectangle();
                    }
                }

                if (picture.BlipFill != null && picture.BlipFill.SourceRectangle != null) {
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
                return picture?.BlipFill?.SourceRectangle?.Left?.Value;
            }
            set {
                _cropLeft = value;
                var picture = GetPicture();
                if (picture == null) return;

                if (picture.BlipFill?.SourceRectangle == null && value != null) {
                    if (picture.BlipFill == null) {
                        picture.BlipFill = new Pic.BlipFill();
                    }
                    if (picture.BlipFill.SourceRectangle == null) {
                        picture.BlipFill.SourceRectangle = new A.SourceRectangle();
                    }
                }

                if (picture.BlipFill != null && picture.BlipFill.SourceRectangle != null) {
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
                return picture?.BlipFill?.SourceRectangle?.Right?.Value;
            }
            set {
                _cropRight = value;
                var picture = GetPicture();
                if (picture == null) return;

                if (picture.BlipFill?.SourceRectangle == null && value != null) {
                    if (picture.BlipFill == null) {
                        picture.BlipFill = new Pic.BlipFill();
                    }
                    if (picture.BlipFill.SourceRectangle == null) {
                        picture.BlipFill.SourceRectangle = new A.SourceRectangle();
                    }
                }

                if (picture.BlipFill != null && picture.BlipFill.SourceRectangle != null) {
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
                    var amount = alpha?.Amount?.Value;
                    if (amount != null) {
                        return (int)((100000 - amount.Value) / 1000);
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
            get {
                if (_Image.Anchor != null) {
                    return WordWrapTextImage.GetWrapTextImage(_Image.Anchor, _Image.Inline ?? new Inline());
                }
                if (_Image.Inline != null) {
                    return WrapTextImage.InLineWithText;
                }
                return null;
            }
            set {
                if (_Image.Anchor != null) {
                    WordWrapTextImage.SetWrapTextImage(_Image, _Image.Anchor, _Image.Inline ?? new Inline(), value);
                } else if (_Image.Inline != null && value != null) {
                    if (value == WrapTextImage.InLineWithText) {
                        return;
                    }
                    var convertedAnchor = WordTextBox.ConvertInlineToAnchor(_Image.Inline, value.Value);
                    _Image.Append(convertedAnchor);
                    _Image.OfType<Inline>().FirstOrDefault()?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets how the image should fill its bounding box. Default is Stretch.
        /// </summary>
        public ImageFillMode FillMode {
            get {
                var picture = GetPicture();
                var blipFill = picture?.BlipFill;
                if (blipFill != null) {
                    var tile = blipFill.GetFirstChild<Tile>();
                    if (tile != null) {
                        if (tile.Alignment?.Value == RectangleAlignmentValues.Center) {
                            _fillMode = ImageFillMode.Center;
                        } else {
                            _fillMode = ImageFillMode.Tile;
                        }
                    } else {
                        var stretch = blipFill.GetFirstChild<Stretch>();
                        if (stretch != null) {
                            _fillMode = stretch.GetFirstChild<FillRectangle>() == null
                                ? ImageFillMode.Fit
                                : ImageFillMode.Stretch;
                        } else {
                            _fillMode = ImageFillMode.Stretch;
                        }
                    }
                }
                return _fillMode;
            }
            set {
                _fillMode = value;
                var picture = GetPicture();
                if (picture == null) return;

                var blipFill = picture.BlipFill;
                if (blipFill == null) return;
                var tile = blipFill.GetFirstChild<Tile>();
                var stretch = blipFill.GetFirstChild<Stretch>();

                switch (value) {
                    case ImageFillMode.Stretch:
                        tile?.Remove();
                        if (stretch == null) {
                            stretch = new Stretch();
                            blipFill.AppendChild(stretch);
                        }
                        if (stretch.GetFirstChild<FillRectangle>() == null) {
                            stretch.AppendChild(new FillRectangle());
                        }
                        break;
                    case ImageFillMode.Tile:
                        stretch?.Remove();
                        if (tile == null) {
                            tile = new Tile();
                            blipFill.AppendChild(tile);
                        }
                        tile.Alignment = null;
                        break;
                    case ImageFillMode.Fit:
                        tile?.Remove();
                        if (stretch == null) {
                            stretch = new Stretch();
                            blipFill.AppendChild(stretch);
                        }
                        var fillRect = stretch.GetFirstChild<FillRectangle>();
                        fillRect?.Remove();
                        break;
                    case ImageFillMode.Center:
                        stretch?.Remove();
                        if (tile == null) {
                            tile = new Tile();
                            blipFill.AppendChild(tile);
                        }
                        tile.Alignment = RectangleAlignmentValues.Center;
                        break;
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
                _useLocalDpi = useLocalDpi?.Val?.Value;
                return _useLocalDpi;
            }
            set {
                _useLocalDpi = value;
                var blip = GetBlip();
                if (blip == null) return;

                var extList = blip.GetFirstChild<BlipExtensionList>();
                if (extList == null) {
                    if (value == null) {
                        return;
                    }
                    extList = new BlipExtensionList();
                    blip.Append(extList);
                }

                var extension = extList
                    .OfType<BlipExtension>()
                    .FirstOrDefault(e => e.Uri == "{28A0092B-C50C-407E-A947-70E740481C1C}");
                if (extension == null) {
                    if (value == null) {
                        return;
                    }
                    extension = new BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
                    extList.Append(extension);
                }

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
                    var hOffset = hPos?.PositionOffset;
                    if (hOffset?.Text != null) {
                        int.TryParse(hOffset.Text, out x);
                    }
                    var vOffset = vPos?.PositionOffset;
                    if (vOffset?.Text != null) {
                        int.TryParse(vOffset.Text, out y);
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
                _preferRelativeResize = nv?.PreferRelativeResize?.Value;
                return _preferRelativeResize;
            }
            set {
                _preferRelativeResize = value;
                var pic = GetPicture();
                if (pic == null) return;
                var nvp = pic.NonVisualPictureProperties ?? (pic.NonVisualPictureProperties = new Pic.NonVisualPictureProperties());
                var nv = nvp.NonVisualPictureDrawingProperties;
                if (nv == null) {
                    if (value == null) {
                        return;
                    }
                    nv = new Pic.NonVisualPictureDrawingProperties();
                    nvp.Append(nv);
                }
                nv.PreferRelativeResize = value;
            }
        }

        /// <summary>
        /// Gets or sets whether the aspect ratio is locked.
        /// </summary>
        public bool? NoChangeAspect {
            get {
                var locks = GetPicture()?.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.PictureLocks;
                _noChangeAspect = locks?.NoChangeAspect?.Value;
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
                _noCrop = locks?.NoCrop?.Value;
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
                _noMove = locks?.NoMove?.Value;
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
                _noResize = locks?.NoResize?.Value;
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
                _noRotation = locks?.NoRotation?.Value;
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
                _noSelection = locks?.NoSelection?.Value;
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
            var nvp = pic.NonVisualPictureProperties ?? (pic.NonVisualPictureProperties = new Pic.NonVisualPictureProperties());
            var nv = nvp.NonVisualPictureDrawingProperties;
            if (nv == null) {
                nv = new Pic.NonVisualPictureDrawingProperties();
                nvp.Append(nv);
            }
            var locks = nv.PictureLocks;
            if (locks == null) {
                locks = new A.PictureLocks();
                nv.Append(locks);
            }
            setter(locks);
        }

    }
}
