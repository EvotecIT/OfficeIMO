using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    // Helpers for inserting images anchored to worksheet cells.
    public partial class ExcelSheet {
        /// <summary>
        /// Enumerates worksheet images anchored in the drawing layer.
        /// </summary>
        public IEnumerable<ExcelImage> Images {
            get {
                var drawingPart = _worksheetPart.DrawingsPart;
                if (drawingPart?.WorksheetDrawing == null) return Enumerable.Empty<ExcelImage>();

                return drawingPart.WorksheetDrawing
                    .ChildElements
                    .Where(IsSupportedImageAnchor)
                    .SelectMany(anchor => anchor.Descendants<Xdr.Picture>().Select(picture => new ExcelImage(picture, anchor, drawingPart, this)))
                    .ToList();
            }
        }

        /// <summary>
        /// Returns an image by non-visual drawing name, or null if it was not found.
        /// </summary>
        public ExcelImage? GetImage(string name) {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return Images.FirstOrDefault(image => string.Equals(image.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Adds an image anchored to the specified cell and returns a wrapper for setting metadata and sizing.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="imageBytes">Image bytes.</param>
        /// <param name="contentType">Content type, e.g. image/png or image/jpeg.</param>
        /// <param name="width">Width as a document length.</param>
        /// <param name="height">Height as a document length.</param>
        /// <param name="offsetX">Optional X offset from cell origin as a document length.</param>
        /// <param name="offsetY">Optional Y offset from cell origin as a document length.</param>
        /// <param name="name">Optional drawing name.</param>
        /// <param name="altText">Optional alternative text description.</param>
        /// <param name="lockAspectRatio">Whether Excel should keep the picture aspect ratio locked.</param>
        public ExcelImage AddImage(int row, int column, byte[] imageBytes, string contentType = "image/png",
            OfficeLength? width = null, OfficeLength? height = null, OfficeLength offsetX = default, OfficeLength offsetY = default,
            string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            if (row <= 0 || column <= 0) throw new ArgumentOutOfRangeException("Row and column are 1-based and must be positive.");
            OfficeLength resolvedWidth = width ?? OfficeLength.Pixels(96);
            OfficeLength resolvedHeight = height ?? OfficeLength.Pixels(32);
            if (resolvedWidth <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (resolvedHeight <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            ExcelImage? image = null;
            WriteLock(() => {
                // Ensure a drawing part exists and is referenced by the worksheet
                DrawingsPart drawingPart;
                var drawing = WorksheetRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
                if (drawing == null) {
                    drawingPart = _worksheetPart.AddNewPart<DrawingsPart>();
                    drawingPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                    string relId = _worksheetPart.GetIdOfPart(drawingPart);
                    WorksheetRoot.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = relId });
                } else {
                    drawingPart = (DrawingsPart)_worksheetPart.GetPartById(drawing.Id!);
                    drawingPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
                }

                // Add the image part
                PartTypeInfo type = ToImagePartType(contentType);
                var imagePart = drawingPart.AddImagePart(type);
                using (var s = new MemoryStream(imageBytes)) imagePart.FeedData(s);
                string imgRelId = drawingPart.GetIdOfPart(imagePart);

                // Build a OneCellAnchor: FromMarker + Extent + Picture + ClientData
                long cx = resolvedWidth.ToEmus();
                long cy = resolvedHeight.ToEmus();
                long dx = offsetX.ToEmus();
                long dy = offsetY.ToEmus();

                var nvId = NextDrawingId(drawingPart);
                string resolvedName = string.IsNullOrWhiteSpace(name) ? $"Picture {nvId}" : name!.Trim();
                var anchor = new Xdr.OneCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId((column - 1).ToString()),
                        new Xdr.ColumnOffset(dx.ToString()),
                        new Xdr.RowId((row - 1).ToString()),
                        new Xdr.RowOffset(dy.ToString())
                    ),
                    new Xdr.Extent { Cx = cx, Cy = cy },
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = nvId, Name = resolvedName, Description = altText ?? string.Empty },
                            new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = lockAspectRatio })
                        ),
                        new Xdr.BlipFill(
                            new A.Blip { Embed = imgRelId },
                            new A.Stretch(new A.FillRectangle())
                        ),
                        new Xdr.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = cx, Cy = cy }
                            ),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                        )
                    ),
                    new Xdr.ClientData()
                );

                drawingPart.WorksheetDrawing.Append(anchor);
                drawingPart.WorksheetDrawing.Save();
                WorksheetRoot.Save();
                _excelDocument.MarkPackageDirty();
                image = new ExcelImage(anchor.GetFirstChild<Xdr.Picture>()!, anchor, drawingPart, this);
            });

            return image!;
        }

        /// <summary>
        /// Asynchronously downloads an image with an explicit remote-network policy and returns its wrapper.
        /// </summary>
        public Task<ExcelImage> AddImageFromUrlWithOptionsAsync(
            int row,
            int column,
            string url,
            OfficeRemoteImageLoadOptions remoteImageOptions,
            OfficeLength? width = null,
            OfficeLength? height = null,
            OfficeLength offsetX = default,
            OfficeLength offsetY = default,
            string? name = null,
            string? altText = null,
            bool lockAspectRatio = true,
            CancellationToken cancellationToken = default) {
            if (remoteImageOptions == null) throw new ArgumentNullException(nameof(remoteImageOptions));
            return AddImageFromUrlPlacementCoreAsync(row, column, url, remoteImageOptions, width, height, null,
                offsetX, offsetY, name, altText, null, lockAspectRatio, 0, cancellationToken);
        }

        private static long PxToEmu(int px) => (long)Math.Round(px * 9525.0);

        private static bool IsSupportedImageAnchor(OpenXmlElement anchor)
            => (anchor is Xdr.OneCellAnchor || anchor is Xdr.TwoCellAnchor || anchor is Xdr.AbsoluteAnchor)
                && anchor.Descendants<Xdr.Picture>().Any();

        private static UInt32Value NextDrawingId(DrawingsPart dp) {
            uint max = 0;
            if (dp.WorksheetDrawing != null) {
                foreach (var nv in dp.WorksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>()) {
                    if (nv.Id != null && nv.Id.Value > max) max = nv.Id.Value;
                }
            }
            return max + 1;
        }
    }
}
