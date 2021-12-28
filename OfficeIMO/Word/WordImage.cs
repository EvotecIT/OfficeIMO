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

//using NonVisualGraphicFrameDrawingProperties = DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties;
//using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
//using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;
//using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
//using ShapeProperties = DocumentFormat.OpenXml.Drawing.ShapeProperties;


namespace OfficeIMO.Word {
    internal class WordImage {
        public static void InsertAPicture(string document, string fileName) {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(document, true)) {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(fileName, FileMode.Open)) {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId) {
            // Define the reference of the image.
            var element =
                new Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() {Cx = 990000L, Cy = 792000L},
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent() {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() {
                            Id = (UInt32Value) 1U,
                            Name = "Picture 1"
                        },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                            new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() {NoChangeAspect = true}),
                        new DocumentFormat.OpenXml.Drawing.Graphic(
                            new DocumentFormat.OpenXml.Drawing.GraphicData(
                                new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() {
                                            Id = (UInt32Value) 0U,
                                            Name = "New Bitmap Image.jpg"
                                        },
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                        new DocumentFormat.OpenXml.Drawing.Blip(
                                            new DocumentFormat.OpenXml.Drawing.BlipExtensionList(
                                                new DocumentFormat.OpenXml.Drawing.BlipExtension() {
                                                    Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                })
                                        ) {
                                            Embed = relationshipId,
                                            CompressionState =
                                                DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                        },
                                        new DocumentFormat.OpenXml.Drawing.Stretch(
                                            new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                        new DocumentFormat.OpenXml.Drawing.Transform2D(
                                            new DocumentFormat.OpenXml.Drawing.Offset() {X = 0L, Y = 0L},
                                            new DocumentFormat.OpenXml.Drawing.Extents() {Cx = 990000L, Cy = 792000L}),
                                        new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                            new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                        ) {Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle}))
                            ) {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
                    ) {
                        DistanceFromTop = (UInt32Value) 0U,
                        DistanceFromBottom = (UInt32Value) 0U,
                        DistanceFromLeft = (UInt32Value) 0U,
                        DistanceFromRight = (UInt32Value) 0U,
                        EditId = "50D07946"
                    });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(element)));
        }
    }
}