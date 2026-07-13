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
    /// <summary>
    /// Helpers for inserting images anchored to worksheet cells.
    /// </summary>
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
                    .SelectMany(anchor => anchor.Descendants<Xdr.Picture>().Select(picture => new ExcelImage(picture, anchor, drawingPart, _excelDocument)))
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
        /// Adds an image anchored to the specified cell. The top-left of the image will align to the cell's top-left,
        /// with optional pixel offsets. Size is specified in pixels and converted to EMUs.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="imageBytes">Image bytes.</param>
        /// <param name="contentType">Content type, e.g. image/png or image/jpeg.</param>
        /// <param name="widthPixels">Width in pixels.</param>
        /// <param name="heightPixels">Height in pixels.</param>
        /// <param name="offsetXPixels">Optional X offset from cell origin in pixels.</param>
        /// <param name="offsetYPixels">Optional Y offset from cell origin in pixels.</param>
        public void AddImageAt(int row, int column, byte[] imageBytes, string contentType = "image/png", int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0) {
            AddImage(row, column, imageBytes, contentType, widthPixels, heightPixels, offsetXPixels, offsetYPixels);
        }

        /// <summary>
        /// Adds an image anchored to the specified cell and returns a wrapper for setting metadata and sizing.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="imageBytes">Image bytes.</param>
        /// <param name="contentType">Content type, e.g. image/png or image/jpeg.</param>
        /// <param name="widthPixels">Width in pixels.</param>
        /// <param name="heightPixels">Height in pixels.</param>
        /// <param name="offsetXPixels">Optional X offset from cell origin in pixels.</param>
        /// <param name="offsetYPixels">Optional Y offset from cell origin in pixels.</param>
        /// <param name="name">Optional drawing name.</param>
        /// <param name="altText">Optional alternative text description.</param>
        /// <param name="lockAspectRatio">Whether Excel should keep the picture aspect ratio locked.</param>
        public ExcelImage AddImage(int row, int column, byte[] imageBytes, string contentType = "image/png",
            int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0,
            string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (imageBytes == null || imageBytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(imageBytes));
            if (row <= 0 || column <= 0) throw new ArgumentOutOfRangeException("Row and column are 1-based and must be positive.");
            if (widthPixels <= 0) throw new ArgumentOutOfRangeException(nameof(widthPixels));
            if (heightPixels <= 0) throw new ArgumentOutOfRangeException(nameof(heightPixels));

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
                long cx = PxToEmu(widthPixels);
                long cy = PxToEmu(heightPixels);
                long dx = PxToEmu(offsetXPixels);
                long dy = PxToEmu(offsetYPixels);

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
                image = new ExcelImage(anchor.GetFirstChild<Xdr.Picture>()!, anchor, drawingPart, _excelDocument);
            });

            return image!;
        }

        /// <summary>
        /// Asynchronously downloads an image from a URL and anchors it to the specified cell.
        /// </summary>
        /// <param name="row">1-based row index where the top edge of the image will be anchored.</param>
        /// <param name="column">1-based column index where the left edge of the image will be anchored.</param>
        /// <param name="url">Remote HTTP or HTTPS image URL.</param>
        /// <param name="widthPixels">Desired image width in pixels. Defaults to 96 px, converted to English Metric Units (EMUs) for OpenXML positioning.</param>
        /// <param name="heightPixels">Desired image height in pixels. Defaults to 32 px, converted to EMUs.</param>
        /// <param name="offsetXPixels">Optional horizontal offset in pixels from the cell's left boundary. Positive values move the image right; defaults to 0 px.</param>
        /// <param name="offsetYPixels">Optional vertical offset in pixels from the cell's top boundary. Positive values move the image down; defaults to 0 px.</param>
        /// <param name="cancellationToken">Token used to cancel the remote request.</param>
        public Task<ExcelImage> AddImageFromUrlAtAsync(int row, int column, string url, int widthPixels = 96, int heightPixels = 32,
            int offsetXPixels = 0, int offsetYPixels = 0, CancellationToken cancellationToken = default)
            => AddImageFromUrlAsync(row, column, url, widthPixels, heightPixels, offsetXPixels, offsetYPixels,
                cancellationToken: cancellationToken);

        /// <summary>
        /// Asynchronously downloads an image from a URL and returns a wrapper for setting metadata and sizing.
        /// </summary>
        public async Task<ExcelImage> AddImageFromUrlAsync(int row, int column, string url, int widthPixels = 96, int heightPixels = 32,
            int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null, bool lockAspectRatio = true,
            CancellationToken cancellationToken = default) {
            OfficeRemoteImage remote = await OfficeRemoteImageLoader.LoadAsync(
                url,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            byte[] bytes = remote.ToBytes();
            OfficeImageReader.TryIdentify(bytes, remote.FileName, out OfficeImageInfo info);
            return AddImage(row, column, bytes, contentType: ResolveImageContentType(remote.ContentType, info), widthPixels: widthPixels,
                heightPixels: heightPixels, offsetXPixels: offsetXPixels, offsetYPixels: offsetYPixels, name: name, altText: altText,
                lockAspectRatio: lockAspectRatio);
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
