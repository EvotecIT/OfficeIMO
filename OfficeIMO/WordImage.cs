using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

//using NonVisualGraphicFrameDrawingProperties = DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties;
//using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
//using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;
//using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
//using ShapeProperties = DocumentFormat.OpenXml.Drawing.ShapeProperties;


namespace OfficeIMO {
    public class WordImage {
        internal Drawing _Image;

        public WordImage() {

        }


        public WordImage(WordDocument document, string filePath, ShapeTypeValues shape = ShapeTypeValues.Rectangle, BlipCompressionValues compressionQuality = BlipCompressionValues.Print ) {
            // Size - https://stackoverflow.com/questions/8082980/inserting-image-into-docx-using-openxml-and-setting-the-size


            
            var fileName = System.IO.Path.GetFileName(filePath);
            var imageName = System.IO.Path.GetFileNameWithoutExtension(filePath);

            ImagePart imagePart = document._wordprocessingDocument.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
            
            using (FileStream stream = new FileStream(filePath, FileMode.Open)) {
                imagePart.FeedData(stream);
            }

            var relationshipId = document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(imagePart);
            
            var shapeProperties = new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                new Transform2D(
                    new Offset() { X = 0L, Y = 0L }, 
                    new Extents() { Cx = 990000L, Cy = 792000L }),
                new PresetGeometry(new AdjustValueList()) { Preset = shape });

            var element =
                new Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 990000L, Cy = 792000L },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent() {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() {
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
                                                })                                        ) {
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
            this._Image = element;
        }
    }
}