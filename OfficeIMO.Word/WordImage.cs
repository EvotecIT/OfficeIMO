using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.Linq;
using Anchor = DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties;


namespace OfficeIMO.Word {
    public class WordImage {
        double englishMetricUnitsPerInch = 914400;
        double pixelsPerInch = 96;

        internal Drawing _Image;
        internal ImagePart _imagePart;

        private readonly WordDocument _document;
        //internal ShapeProperties _shapeProperties;

        public BlipCompressionValues? CompressionQuality {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    return picture.BlipFill.Blip.CompressionState;
                }
                // TODO: we need to take care of non INLINE images
                return null;
            }
            set { }
        }

        public string RelationshipId {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    return picture.BlipFill.Blip.Embed;
                }
                // TODO: we need to take care of non INLINE images
                return null;
            }
            set {

            }
        }

        public string FilePath { get; set; }

        public string FileName {
            get {
                if (_Image.Inline != null) {
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    return picture.NonVisualPictureProperties.NonVisualDrawingProperties.Name;
                }
                // TODO: we need to take care of non INLINE images
                return null;
            }
            set {

            }
        }

        public double? Width {
            get {
                if (_Image.Inline != null) {
                    var extents = _Image.Inline.Extent;
                    var cX = extents.Cx;
                    return cX / englishMetricUnitsPerInch * pixelsPerInch;
                }
                // TODO: we need to take care of non INLINE images
                return null;
            }
            set {
                double emuWidth = value.Value * englishMetricUnitsPerInch / pixelsPerInch;
                _Image.Inline.Extent.Cx = (Int64Value)emuWidth;
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                //var picture = _Image.Inline.Graphic.GraphicData.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().First();
                picture.ShapeProperties.Transform2D.Extents.Cx = (Int64Value)emuWidth;
            }
        }
        public double? Height {
            get {
                if (_Image.Inline != null) {
                    var extents = _Image.Inline.Extent;
                    var cY = extents.Cy;
                    return cY / englishMetricUnitsPerInch * pixelsPerInch;
                }
                // TODO: we need to take care of non INLINE images
                return null;
            }
            set {
                if (_Image.Inline != null) {
                    double emuHeight = value.Value * englishMetricUnitsPerInch / pixelsPerInch;
                    _Image.Inline.Extent.Cy = (Int64Value)emuHeight;
                    var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                    //var picture = _Image.Inline.Graphic.GraphicData.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().First();
                    picture.ShapeProperties.Transform2D.Extents.Cy = (Int64Value)emuHeight;
                } else {
                    Console.WriteLine("Not supported yet");
                }
            }
        }
        public double? EmuWidth {
            get {
                if (_Image.Inline != null) {
                    var extents = _Image.Inline.Extent;
                    return extents.Cx;
                }
                // TODO: we need to take care of non INLINE images
                return null;
            }
        }
        public double? EmuHeight {
            get {
                if (_Image.Inline != null) {
                    var extents = _Image.Inline.Extent;
                    return extents.Cy;
                }
                // TODO: we need to take care of non INLINE images
                return null;
            }
        }

        public ShapeTypeValues Shape {
            get {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                var presetGeometry = picture.ShapeProperties.GetFirstChild<PresetGeometry>();
                return presetGeometry.Preset;
            }
            set {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                var presetGeometry = picture.ShapeProperties.GetFirstChild<PresetGeometry>();
                presetGeometry.Preset = value;
            }
        }
        public BlackWhiteModeValues BlackWiteMode {
            // this doesn't work
            get {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                return picture.ShapeProperties.BlackWhiteMode.Value;
            }
            set {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                if (picture.ShapeProperties.BlackWhiteMode == null) {
                    picture.ShapeProperties.BlackWhiteMode = new EnumValue<BlackWhiteModeValues>();
                }
                picture.ShapeProperties.BlackWhiteMode.Value = value;
            }
        }
        public bool VerticalFlip {
            // this doesn't work
            get {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                return picture.ShapeProperties.Transform2D.VerticalFlip;
            }
            set {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                picture.ShapeProperties.Transform2D.VerticalFlip = value;
            }
        }
        public bool HorizontalFlip {
            // this doesn't work
            get {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                return picture.ShapeProperties.Transform2D.HorizontalFlip;
            }
            set {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                picture.ShapeProperties.Transform2D.HorizontalFlip = value;
            }
        }
        public int Rotation {
            // this doesn't work
            get {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                return picture.ShapeProperties.Transform2D.Rotation;
            }
            set {
                var picture = _Image.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
                picture.ShapeProperties.Transform2D.Rotation = value;
            }
        }

        public bool? Wrap {
            // this doesn't work
            get {

                //_Image.Anchor.
                return null;
            }
            set {

                if (_Image.Anchor == null) {
                    var inline = _Image.Inline.CloneNode(true);

                    IEnumerable<OpenXmlElement> clonedElements = _Image.Inline
                        .Elements()
                        .Select(e => e.CloneNode(true))
                        .ToList();

                    var childInline = inline.Descendants();
                    Anchor anchor1 = new Anchor() { BehindDoc = true };
                    WrapNone wrapNone1 = new WrapNone();
                    anchor1.Append(wrapNone1);
                    _Image.Append(anchor1);

                    _Image.Inline.Remove();

                    _Image.Anchor.Append(clonedElements);
                } else {
                    _Image.Anchor.AllowOverlap = true;
                }
            }
        }



        //public WordImage(WordDocument document, WordParagraph paragraph, string filePath, ShapeTypeValues shape = ShapeTypeValues.Rectangle, BlipCompressionValues compressionQuality = BlipCompressionValues.Print) {
        //    double width;
        //    double height;
        //    using (var img = SixLabors.ImageSharp.Image.Load(filePath)) {
        //        width = img.Width;
        //        height = img.Height;
        //    }
        //}

        //public WordImage(WordDocument document, Paragraph paragraph) {

        //}

        public WordImage(WordDocument document, WordParagraph paragraph, string filePath, double? width, double? height, ShapeTypeValues shape = ShapeTypeValues.Rectangle, BlipCompressionValues compressionQuality = BlipCompressionValues.Print) {
            _document = document;

            // Size - https://stackoverflow.com/questions/8082980/inserting-image-into-docx-using-openxml-and-setting-the-size

            //var uri = new Uri(filePath, UriKind.RelativeOrAbsolute);
            using var imageStream = new FileStream(filePath, FileMode.Open);

            // if widht/height are not set we check ourselves
            // but probably will need better way
            var imageСharacteristics = Helpers.GetImageСharacteristics(imageStream);
            if (width == null || height == null) {
                width = imageСharacteristics.Width;
                height = imageСharacteristics.Height;
            }

            var fileName = System.IO.Path.GetFileName(filePath);
            var imageName = System.IO.Path.GetFileNameWithoutExtension(filePath);

            // decide where to put an image based on the location of paragraph
            ImagePart imagePart;
            string relationshipId;
            var location = paragraph.Location();
            if (location.GetType() == typeof(Header)) {
                var part = ((Header)location).HeaderPart;
                imagePart = part.AddImagePart(imageСharacteristics.Type);
                relationshipId = part.GetIdOfPart(imagePart);
            } else if (location.GetType() == typeof(Footer)) {
                var part = ((Footer)location).FooterPart;
                imagePart = part.AddImagePart(imageСharacteristics.Type);
                relationshipId = part.GetIdOfPart(imagePart);
            } else if (location.GetType() == typeof(Document)) {
                var part = document._wordprocessingDocument.MainDocumentPart;
                imagePart = part.AddImagePart(imageСharacteristics.Type);
                relationshipId = part.GetIdOfPart(imagePart);
            } else {
                throw new Exception("Paragraph is not in document or header or footer. This is weird. Probably a bug.");
            }

            this._imagePart = imagePart;
            imagePart.FeedData(imageStream);

            //calculate size in emu
            double emuWidth = width.Value * englishMetricUnitsPerInch / pixelsPerInch;
            double emuHeight = height.Value * englishMetricUnitsPerInch / pixelsPerInch;

            var shapeProperties = new ShapeProperties(
                new Transform2D(new Offset() { X = 0L, Y = 0L },
                    new Extents() {
                        Cx = (Int64Value)emuWidth,
                        Cy = (Int64Value)emuHeight
                    }
                    ),
                new PresetGeometry(new AdjustValueList()) { Preset = shape }
                );

            //this._shapeProperties = shapeProperties;

            var drawing = new Drawing(

            //new Anchor(
            //    new WrapNone()
            //    ) { BehindDoc = true },

            new Inline(
                        new Extent() { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight },

                        new EffectExtent() {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DocProperties() {
                            Id = (UInt32Value)1U,
                            Name = imageName
                        },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                            new GraphicFrameLocks() { NoChangeAspect = true }),
                        new Graphic(
                            new GraphicData(
                                new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() {
                                            Id = (UInt32Value)0U,
                                            Name = fileName
                                        },
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                        new Blip(new BlipExtensionList(new BlipExtension() {
                                            // https://stackoverflow.com/questions/33521914/value-of-blipextension-schema-uri-28a0092b-c50c-407e-a947-70e740481c1c
                                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                        })
                                        ) {
                                            Embed = relationshipId,
                                            CompressionState = compressionQuality
                                        },
                                        new Stretch(new FillRectangle())),
                                  shapeProperties)
                            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    ) {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946"
            });

            this._Image = drawing;
            //this.Width = width.Value;
            //this.Height = height.Value;
            //this.EmuWidth = emuWidth;
            //this.EmuHeight = emuHeight;
            this.Shape = shape;
            //this.CompressionQuality = compressionQuality;
            //this.FileName = fileName;
            this.FilePath = filePath;
            //this.RelationshipId = relationshipId;

            // document.Images.Add(this);
        }

        public WordImage(WordDocument document, Drawing drawing) {
            _document = document;
            _Image = drawing;
            var imageParts = document._document.MainDocumentPart.ImageParts;
            foreach (var imagePart in imageParts) {
                var relationshipId = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(imagePart);
                if (this.RelationshipId == relationshipId) {
                    this._imagePart = imagePart;
                }
            }
        }

        /// <summary>
        /// Extract image from Word Document and save it to file
        /// </summary>
        /// <param name="fileToSave"></param>
        public void SaveToFile(string fileToSave) {
            using (FileStream outputFileStream = new FileStream(fileToSave, FileMode.Create)) {
                var stream = this._imagePart.GetStream();
                stream.CopyTo(outputFileStream);
                stream.Close();
            }
        }

        public void Remove() {
            if (this._imagePart != null) {
                _document._wordprocessingDocument.MainDocumentPart.DeletePart(_imagePart);
            }

            if (this._Image != null) {
                this._Image.Remove();
            }
        }
    }
}
