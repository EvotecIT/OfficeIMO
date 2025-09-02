using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp;
using System;
using System.IO;
using System.Net.Http;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides methods for embedding images into Wordprocessing documents.
    /// </summary>
    public static class ImageEmbedder {
        /// <summary>
        /// Creates a run element containing an image loaded from the specified source.
        /// </summary>
        /// <param name="mainPart">Main document part where the image will be stored.</param>
        /// <param name="src">Path, URI or data URI representing the image.</param>
        /// <returns>A <see cref="Run"/> containing the embedded image.</returns>
        public static Run CreateImageRun(MainDocumentPart mainPart, string src) {
            byte[] bytes = ResolveImageSource(src);
            using Image image = Image.Load(bytes, out var format);
            long cx = (long)(image.Width * 9525L);
            long cy = (long)(image.Height * 9525L);
            string contentType = format.DefaultMimeType;

            ImagePart imagePart = mainPart.AddImagePart(contentType);
            using (MemoryStream ms = new MemoryStream(bytes)) {
                imagePart.FeedData(ms);
            }
            string relationshipId = mainPart.GetIdOfPart(imagePart);

            var inline = new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = 1U, Name = "Picture" },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "Image" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            ) { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U };

            var drawing = new Drawing(inline);
            return new Run(drawing);
        }

        /// <summary>
        /// Retrieves raw bytes from a <see cref="WordImage"/> instance.
        /// </summary>
        /// <param name="image">Image to extract bytes from.</param>
        /// <returns>Binary data representing the image.</returns>
        public static byte[] GetImageBytes(WordImage image) {
            return image.GetBytes();
        }

        private static byte[] ResolveImageSource(string src) {
            if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
                int commaIndex = src.IndexOf(',');
                string base64Data = src.Substring(commaIndex + 1);
                return Convert.FromBase64String(base64Data);
            }

            if (Uri.TryCreate(src, UriKind.Absolute, out var uri)) {
                if (uri != null) {
                    if (uri.Scheme == Uri.UriSchemeFile) {
                        return File.ReadAllBytes(uri.LocalPath);
                    }
                    if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps) {
                        using HttpClient client = new HttpClient();
                        return client.GetByteArrayAsync(uri).GetAwaiter().GetResult();
                    }
                }
            }

            if (File.Exists(src)) {
                return File.ReadAllBytes(src);
            }

            throw new InvalidOperationException("Unable to resolve image source: " + src);
        }
    }
}
