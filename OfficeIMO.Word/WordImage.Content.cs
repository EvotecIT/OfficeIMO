using DocumentFormat.OpenXml;
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
        /// Wraps an existing WordDrawing as a WordImage.
        /// </summary>
        public WordImage(WordDocument document, WordDrawing drawing) {
            _document = document;
            _Image = drawing;

            var initialBlip = GetBlip();
            if (initialBlip != null) {
                if (initialBlip.Link != null) {
                    _externalRelationshipId = initialBlip.Link;
                } else if (initialBlip.Embed?.Value is { Length: > 0 } embedId) {
                    var part = GetContainingPart();
                    _imagePart = part.GetPartById(embedId) as ImagePart;
                }
            }

            var picture = GetPicture();
            if (picture != null) {
                var nv = picture.NonVisualPictureProperties;
                if (nv != null) {
                    _title = nv.NonVisualDrawingProperties?.Title;
                    _hidden = nv.NonVisualDrawingProperties?.Hidden?.Value;
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
                    if (ar?.Alpha != null) _fixedOpacity = (int?)(ar.Alpha.Value / 1000);
                    var ai = pictureBlip.GetFirstChild<AlphaInverse>();
                    _alphaInversionColorHex = ai?.GetFirstChild<RgbColorModelHex>()?.Val;
                    var bi = pictureBlip.GetFirstChild<BiLevel>();
                    _blackWhiteThreshold = bi?.Threshold?.Value != null ? (int?)(bi.Threshold.Value / 1000) : null;
                    var blur = pictureBlip.GetFirstChild<Blur>();
                    if (blur != null) { _blurRadius = (int?)blur.Radius?.Value; _blurGrow = blur.Grow?.Value; }
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
        /// Wraps an existing VML image as a WordImage.
        /// </summary>
        internal WordImage(WordDocument document, DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph, DocumentFormat.OpenXml.Wordprocessing.Run run, V.Shape shape) {
            _document = document;
            _vmlShape = shape;
            _vmlImageData = shape.GetFirstChild<V.ImageData>();
        }

        /// <summary>
        /// Creates a copy of this image and appends it to the specified paragraph.
        /// The cloned image shares the same underlying image part.
        /// </summary>
        /// <param name="paragraph">The paragraph to append the cloned image to.</param>
        /// <returns>The newly created <see cref="WordImage"/> instance.</returns>
        public WordImage Clone(WordParagraph paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));

            var drawingClone = (WordDrawing)_Image.CloneNode(true);
            var run = new DocumentFormat.OpenXml.Wordprocessing.Run(drawingClone);
            paragraph._paragraph.Append(run);

            return new WordImage(paragraph._document, drawingClone);
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
                using (FileStream outputFileStream = new FileStream(fileToSave, FileMode.Create, FileAccess.Write, FileShare.None)) {
                    using var stream = _imagePart.GetStream(FileMode.Open, FileAccess.Read);
                    stream.CopyTo(outputFileStream);
                }
            } catch (UnauthorizedAccessException ex) {
                throw new IOException($"Failed to save to '{fileToSave}'. Access denied or path is read-only.", ex);
            }
        }

        /// <summary>
        /// Retrieves the image data as a stream without loading the entire image into memory.
        /// </summary>
        /// <returns>A <see cref="Stream"/> for reading the image bytes.</returns>
        public Stream GetStream() {
            if (_imagePart == null) {
                throw new InvalidOperationException("Image is linked externally and cannot be extracted.");
            }

            return _imagePart.GetStream(FileMode.Open, FileAccess.Read);
        }

        /// <summary>
        /// Retrieves the image data as a byte array.
        /// </summary>
        /// <returns>Bytes representing the image.</returns>
        public byte[] GetBytes() {
            using var stream = GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }

        /// <summary>
        /// Remove image from a Word Document
        /// </summary>
        public void Remove() {
            if (_imagePart != null) {
                OpenXmlElement? parent = _Image?.Parent;
                while (parent != null && parent is not Body && parent is not Header && parent is not Footer) {
                    parent = parent.Parent;
                }

                OpenXmlPart? part = _document._wordprocessingDocument.MainDocumentPart;
                if (parent is Header header) {
                    part = header.HeaderPart;
                } else if (parent is Footer footer) {
                    part = footer.FooterPart;
                }

                part?.DeletePart(_imagePart);
                _imagePart = null;
            } else if (!string.IsNullOrEmpty(_externalRelationshipId)) {
                OpenXmlElement? parent = _Image?.Parent;
                while (parent != null && parent is not Body && parent is not Header && parent is not Footer) {
                    parent = parent.Parent;
                }

                OpenXmlPart? part = _document._wordprocessingDocument.MainDocumentPart;
                if (parent is Header header) {
                    part = header.HeaderPart;
                } else if (parent is Footer footer) {
                    part = footer.FooterPart;
                }

                if (part != null) {
                    var rel = part.ExternalRelationships.FirstOrDefault(r => r.Id == _externalRelationshipId);
                    if (rel != null) {
                        part.DeleteExternalRelationship(rel);
                    }
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

            var drawing = new WordDrawing();
            UInt32Value docPropertiesId = GetNextDocPropertiesId(document);

            if (wrapImage == WrapTextImage.InLineWithText) {
                var inline = GetInline(emuWidth, emuHeight, docPropertiesId, imageLocation.ImageName, fileName, imageLocation.RelationshipId, shape, compressionQuality, description);
                drawing.Append(inline);
            } else {
                var graphic = GetGraphic(emuWidth, emuHeight, fileName, imageLocation.RelationshipId, shape, compressionQuality, description);
                var anchor = GetAnchor(emuWidth, emuHeight, docPropertiesId, graphic, imageLocation.ImageName, description, wrapImage);
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
            return Helpers.UseSeekableImageStream(imageStream, preparedImageStream => {
                // Size - https://stackoverflow.com/questions/8082980/inserting-image-into-docx-using-openxml-and-setting-the-size
                // if widht/height are not set we check ourselves
                // but probably will need better way
                var imageCharacteristics = Helpers.GetImageCharacteristics(preparedImageStream, fileName);
                if (width == null || height == null) {
                    if (imageCharacteristics.Width == 0 || imageCharacteristics.Height == 0) {
                        throw new ArgumentException("Width and height must be provided for this image type.");
                    }

                    if (width != null) {
                        height = imageCharacteristics.Height * width.Value / imageCharacteristics.Width;
                    } else if (height != null) {
                        width = imageCharacteristics.Width * height.Value / imageCharacteristics.Height;
                    } else {
                        width = imageCharacteristics.Width;
                        height = imageCharacteristics.Height;
                    }
                }

                var imagePartType = imageCharacteristics.Type;
                var imageName = System.IO.Path.GetFileNameWithoutExtension(fileName);

                ImagePart imagePart;
                string relationshipId;
                var location = paragraph.Location();
                if (location.GetType() == typeof(Header)) {
                    var part = ((Header)location).HeaderPart ?? throw new InvalidOperationException("Header part is missing.");
                    imagePart = part.AddImagePart(imagePartType.ToOpenXmlImagePartType());
                    relationshipId = part.GetIdOfPart(imagePart);
                } else if (location.GetType() == typeof(Footer)) {
                    var part = ((Footer)location).FooterPart ?? throw new InvalidOperationException("Footer part is missing.");
                    imagePart = part.AddImagePart(imagePartType.ToOpenXmlImagePartType());
                    relationshipId = part.GetIdOfPart(imagePart);
                } else if (location.GetType() == typeof(Document)) {
                    var part = document._wordprocessingDocument.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is missing.");
                    imagePart = part.AddImagePart(imagePartType.ToOpenXmlImagePartType());
                    relationshipId = part.GetIdOfPart(imagePart);
                } else {
                    throw new InvalidOperationException("Paragraph is not in document or header or footer. This is weird. Probably a bug.");
                }

                preparedImageStream.Position = 0;
                imagePart.FeedData(preparedImageStream);

                return new WordImageLocation() {
                    ImagePart = imagePart,
                    RelationshipId = relationshipId,
                    Width = width.Value,
                    Height = height.Value,
                    ImageName = imageName
                };
            });
        }

        private void AddExternalImage(WordDocument document, WordParagraph paragraph, Uri uri, double width, double height, ShapeTypeValues shape, BlipCompressionValues compressionQuality, string description, WrapTextImage wrapImage) {
            _document = document;

            var location = paragraph.Location();
            ExternalRelationship rel;
            if (location is Header header) {
                var part = header.HeaderPart ?? throw new InvalidOperationException("Header part is missing.");
                rel = part.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", uri);
            } else if (location is Footer footer) {
                var part = footer.FooterPart ?? throw new InvalidOperationException("Footer part is missing.");
                rel = part.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", uri);
            } else {
                var part = document._wordprocessingDocument.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is missing.");
                rel = part.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", uri);
            }

            _externalRelationshipId = rel.Id;

            double emuWidth = width * EnglishMetricUnitsPerInch / PixelsPerInch;
            double emuHeight = height * EnglishMetricUnitsPerInch / PixelsPerInch;

            var drawing = new WordDrawing();
            UInt32Value docPropertiesId = GetNextDocPropertiesId(document);
            if (wrapImage == WrapTextImage.InLineWithText) {
                var inline = GetInline(emuWidth, emuHeight, docPropertiesId, System.IO.Path.GetFileNameWithoutExtension(uri.ToString()), System.IO.Path.GetFileName(uri.ToString()), rel.Id, shape, compressionQuality, description, true);
                drawing.Append(inline);
            } else {
                var graphic = GetGraphic(emuWidth, emuHeight, System.IO.Path.GetFileName(uri.ToString()), rel.Id, shape, compressionQuality, description, true);
                var anchor = GetAnchor(emuWidth, emuHeight, docPropertiesId, graphic, System.IO.Path.GetFileNameWithoutExtension(uri.ToString()), description, wrapImage);
                drawing.Append(anchor);
            }
            _Image = drawing;
        }

        private static UInt32Value GetNextDocPropertiesId(WordDocument document) {
            uint max = 0U;
            var mainPart = document._wordprocessingDocument.MainDocumentPart;

            if (mainPart?.Document != null) {
                foreach (DocProperties properties in mainPart.Document.Descendants<DocProperties>()) {
                    if (properties.Id != null && properties.Id.Value > max) {
                        max = properties.Id.Value;
                    }
                }
            }

            if (mainPart != null) {
                foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                    if (headerPart.Header == null) {
                        continue;
                    }

                    foreach (DocProperties properties in headerPart.Header.Descendants<DocProperties>()) {
                        if (properties.Id != null && properties.Id.Value > max) {
                            max = properties.Id.Value;
                        }
                    }
                }

                foreach (FooterPart footerPart in mainPart.FooterParts) {
                    if (footerPart.Footer == null) {
                        continue;
                    }

                    foreach (DocProperties properties in footerPart.Footer.Descendants<DocProperties>()) {
                        if (properties.Id != null && properties.Id.Value > max) {
                            max = properties.Id.Value;
                        }
                    }
                }

                if (mainPart.FootnotesPart?.Footnotes != null) {
                    foreach (DocProperties properties in mainPart.FootnotesPart.Footnotes.Descendants<DocProperties>()) {
                        if (properties.Id != null && properties.Id.Value > max) {
                            max = properties.Id.Value;
                        }
                    }
                }

                if (mainPart.EndnotesPart?.Endnotes != null) {
                    foreach (DocProperties properties in mainPart.EndnotesPart.Endnotes.Descendants<DocProperties>()) {
                        if (properties.Id != null && properties.Id.Value > max) {
                            max = properties.Id.Value;
                        }
                    }
                }
            }

            return (UInt32Value)(max + 1U);
        }

        private OpenXmlPart GetContainingPart() {
            OpenXmlElement? parent = _Image.Parent;
            while (parent != null && parent is not Body && parent is not Header && parent is not Footer) {
                parent = parent.Parent;
            }

            if (parent is Header header) {
                return header.HeaderPart ?? throw new InvalidOperationException("Header part is missing.");
            }

            if (parent is Footer footer) {
                return footer.FooterPart ?? throw new InvalidOperationException("Footer part is missing.");
            }

            return _document._wordprocessingDocument.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is missing.");
        }
    }
}
