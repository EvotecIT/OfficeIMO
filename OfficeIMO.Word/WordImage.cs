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
using DocumentFormat.OpenXml.Office2010.Word.Drawing;

namespace OfficeIMO.Word {
    public class WordImage {
        private const double EnglishMetricUnitsPerInch = 914400;
        private const double PixelsPerInch = 96;

        internal Drawing _Image;
        private ImagePart _imagePart;
        private WordDocument _document;

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

        /// <summary>
        /// Gets or sets the image's wrap text.
        /// </summary>
        public WrapTextImage? WrapText {
            get => WordWrapTextImage.GetWrapTextImage(_Image.Anchor, _Image.Inline);
            set => WordWrapTextImage.SetWrapTextImage(_Image, _Image.Anchor, _Image.Inline, value);
        }

        public WordImage(
            WordDocument document,
            WordParagraph paragraph,
            string filePath,
            double? width,
            double? height,
            WrapTextImage wrapImage = WrapTextImage.InLineWithText,
            string description = "",
            ShapeTypeValues shape = ShapeTypeValues.Rectangle,
            BlipCompressionValues compressionQuality = BlipCompressionValues.Print) {
            FilePath = filePath;
            var fileName = System.IO.Path.GetFileName(filePath);
            using var imageStream = new FileStream(filePath, FileMode.Open);
            AddImage(document, paragraph, imageStream, fileName, width, height, shape, compressionQuality, description, wrapImage);
        }

        public WordImage(
            WordDocument document,
            WordParagraph paragraph,
            Stream imageStream,
            string fileName,
            double? width,
            double? height,
            WrapTextImage wrapImage = WrapTextImage.InLineWithText,
            string description = "",
            ShapeTypeValues shape = ShapeTypeValues.Rectangle,
            BlipCompressionValues compressionQuality = BlipCompressionValues.Print) {
            FilePath = fileName;
            AddImage(document, paragraph, imageStream, fileName, width, height, shape, compressionQuality, description, wrapImage);
        }

        private Graphic GetGraphic(double emuWidth, double emuHeight, string fileName, string relationshipId, ShapeTypeValues shape, BlipCompressionValues compressionQuality, string description = "") {

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

            var nonVisualPictureProperties = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() {
                    Id = (UInt32Value)0U,
                    Name = fileName,
                    // this description doesn't seem to matter, but leaving it here for now
                    Description = description
                },
                //new Pic.NonVisualPictureProperties(),
                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties() { });

            var blipFlip = new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill();

            var blip = new Blip() { Embed = relationshipId, CompressionState = compressionQuality };

            // https://stackoverflow.com/questions/33521914/value-of-blipextension-schema-uri-28a0092b-c50c-407e-a947-70e740481c1c
            var blipExtensionList = new BlipExtensionList();
            blip.Append(blipExtensionList);

            var blipExtension1 = new BlipExtension() {
                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
            };

            DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            blipExtension1.Append(useLocalDpi1);

            blipExtensionList.Append(blipExtension1);

            var stretch = new Stretch(new FillRectangle());

            blipFlip.Append(blip);
            blipFlip.Append(stretch);

            var picture = new DocumentFormat.OpenXml.Drawing.Pictures.Picture();

            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFlip);
            picture.Append(shapeProperties);

            graphic.Append(graphicData);
            graphicData.Append(picture);

            return graphic;
        }

        private Inline GetInline(double emuWidth, double emuHeight, string imageName, string fileName, string relationshipId, ShapeTypeValues shape, BlipCompressionValues compressionQuality, string description = "") {
            var inline = new Inline() {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946"
            };
            inline.Append(new Extent() { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight });
            inline.Append(new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });
            inline.Append(new DocProperties() { Id = (UInt32Value)1U, Name = imageName, Description = description });
            inline.Append(new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new GraphicFrameLocks() { NoChangeAspect = true }));
            inline.Append(GetGraphic(emuWidth, emuHeight, fileName, relationshipId, shape, compressionQuality));

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

            DocProperties docProperties1 = new DocProperties() { Id = (UInt32Value)1U, Name = imageName, Description = description };
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

        /// <summary>
        /// Remove image from a Word Document
        /// </summary>
        public void Remove() {
            if (this._imagePart != null) {
                _document._wordprocessingDocument.MainDocumentPart.DeletePart(_imagePart);
            }

            if (this._Image != null) {
                this._Image.Remove();
            }
        }

        private void AddImage(
            WordDocument document,
            WordParagraph paragraph,
            Stream imageStream,
            string fileName,
            double? width,
            double? height,
            ShapeTypeValues shape,
            BlipCompressionValues compressionQuality,
            string description,
            WrapTextImage wrapImage
        ) {
            _document = document;
            // Size - https://stackoverflow.com/questions/8082980/inserting-image-into-docx-using-openxml-and-setting-the-size
            // if widht/height are not set we check ourselves
            // but probably will need better way
            var imageСharacteristics = Helpers.GetImageСharacteristics(imageStream);
            if (width == null || height == null) {
                width = imageСharacteristics.Width;
                height = imageСharacteristics.Height;
            }

            //var fileName = System.IO.Path.GetFileName(filePath);
            var imageName = System.IO.Path.GetFileNameWithoutExtension(fileName);

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
            double emuWidth = width.Value * EnglishMetricUnitsPerInch / PixelsPerInch;
            double emuHeight = height.Value * EnglishMetricUnitsPerInch / PixelsPerInch;

            var drawing = new Drawing();

            if (wrapImage == WrapTextImage.InLineWithText) {
                var inline = GetInline(emuWidth, emuHeight, imageName, fileName, relationshipId, shape, compressionQuality, description);
                drawing.Append(inline);
            } else {
                var graphic = GetGraphic(emuWidth, emuHeight, fileName, relationshipId, shape, compressionQuality, description);
                var anchor = GetAnchor(emuWidth, emuHeight, graphic, imageName, description, wrapImage);
                drawing.Append(anchor);
            }
            this._Image = drawing;
        }
    }
}
