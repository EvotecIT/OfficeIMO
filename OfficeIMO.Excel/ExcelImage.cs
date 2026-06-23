using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a worksheet image anchored in the drawing layer.
    /// </summary>
    public sealed class ExcelImage {
        private readonly Xdr.Picture _picture;
        private readonly OpenXmlElement _anchor;
        private readonly DrawingsPart _drawingsPart;

        internal ExcelImage(Xdr.Picture picture, OpenXmlElement anchor, DrawingsPart drawingsPart) {
            _picture = picture ?? throw new ArgumentNullException(nameof(picture));
            _anchor = anchor ?? throw new ArgumentNullException(nameof(anchor));
            _drawingsPart = drawingsPart ?? throw new ArgumentNullException(nameof(drawingsPart));
        }

        /// <summary>
        /// Gets or sets the image drawing name.
        /// </summary>
        public string Name {
            get => DrawingProperties?.Name?.Value ?? string.Empty;
            set {
                if (DrawingProperties != null) {
                    DrawingProperties.Name = value ?? string.Empty;
                    Save();
                }
            }
        }

        /// <summary>
        /// Gets or sets the image title metadata.
        /// </summary>
        public string Title {
            get => DrawingProperties?.Title?.Value ?? string.Empty;
            set {
                if (DrawingProperties != null) {
                    DrawingProperties.Title = value ?? string.Empty;
                    Save();
                }
            }
        }

        /// <summary>
        /// Gets or sets the image alternative text description.
        /// </summary>
        public string Description {
            get => DrawingProperties?.Description?.Value ?? string.Empty;
            set {
                if (DrawingProperties != null) {
                    DrawingProperties.Description = value ?? string.Empty;
                    Save();
                }
            }
        }

        /// <summary>
        /// Gets whether Excel should keep the picture aspect ratio locked.
        /// </summary>
        public bool IsAspectRatioLocked {
            get {
                var locks = _picture.NonVisualPictureProperties?.NonVisualPictureDrawingProperties?.GetFirstChild<A.PictureLocks>();
                return locks?.NoChangeAspect?.Value ?? false;
            }
        }

        /// <summary>
        /// Gets the 1-based row index where the image is anchored, when available.
        /// </summary>
        public int RowIndex => GetMarkerRow() + 1;

        /// <summary>
        /// Gets the 1-based column index where the image is anchored, when available.
        /// </summary>
        public int ColumnIndex => GetMarkerColumn() + 1;

        /// <summary>
        /// Gets the image width in pixels from the drawing extent.
        /// </summary>
        public int WidthPixels => EmuToPx(GetExtentCx());

        /// <summary>
        /// Gets the image height in pixels from the drawing extent.
        /// </summary>
        public int HeightPixels => EmuToPx(GetExtentCy());

        /// <summary>
        /// Gets the image horizontal offset from the anchor cell origin in pixels.
        /// </summary>
        public int OffsetXPixels => EmuOffsetToPx(GetMarkerColumnOffset());

        /// <summary>
        /// Gets the image vertical offset from the anchor cell origin in pixels.
        /// </summary>
        public int OffsetYPixels => EmuOffsetToPx(GetMarkerRowOffset());

        /// <summary>
        /// Gets whether the image is positioned with a two-cell drawing anchor.
        /// </summary>
        public bool HasTwoCellAnchor => _anchor is Xdr.TwoCellAnchor;

        /// <summary>
        /// Gets the 1-based ending row index for two-cell anchored images, when available.
        /// </summary>
        public int? ToRowIndex => TryGetToMarkerRow(out int row) ? row + 1 : null;

        /// <summary>
        /// Gets the 1-based ending column index for two-cell anchored images, when available.
        /// </summary>
        public int? ToColumnIndex => TryGetToMarkerColumn(out int column) ? column + 1 : null;

        /// <summary>
        /// Gets the horizontal offset from the two-cell ending marker column in pixels.
        /// </summary>
        public int ToOffsetXPixels => EmuOffsetToPx(GetToMarkerColumnOffset());

        /// <summary>
        /// Gets the vertical offset from the two-cell ending marker row in pixels.
        /// </summary>
        public int ToOffsetYPixels => EmuOffsetToPx(GetToMarkerRowOffset());

        /// <summary>
        /// Gets the left crop ratio authored on the image, where 0 means no crop and 1 means the full width.
        /// </summary>
        public double CropLeftRatio => GetCropRatio(CropRectangle?.Left?.Value);

        /// <summary>
        /// Gets the top crop ratio authored on the image, where 0 means no crop and 1 means the full height.
        /// </summary>
        public double CropTopRatio => GetCropRatio(CropRectangle?.Top?.Value);

        /// <summary>
        /// Gets the right crop ratio authored on the image, where 0 means no crop and 1 means the full width.
        /// </summary>
        public double CropRightRatio => GetCropRatio(CropRectangle?.Right?.Value);

        /// <summary>
        /// Gets the bottom crop ratio authored on the image, where 0 means no crop and 1 means the full height.
        /// </summary>
        public double CropBottomRatio => GetCropRatio(CropRectangle?.Bottom?.Value);

        /// <summary>
        /// Gets the clockwise picture rotation in degrees authored on the image transform.
        /// </summary>
        public double RotationDegrees => GetRotationDegrees(Transform?.Rotation?.Value);

        /// <summary>
        /// Gets whether the image has an authored horizontal flip transform.
        /// </summary>
        public bool FlipHorizontal => Transform?.HorizontalFlip?.Value ?? false;

        /// <summary>
        /// Gets whether the image has an authored vertical flip transform.
        /// </summary>
        public bool FlipVertical => Transform?.VerticalFlip?.Value ?? false;

        /// <summary>
        /// Gets this image anchor's zero-based order in the worksheet drawing layer.
        /// </summary>
        public int DrawingOrder => GetDrawingOrder(_anchor, _drawingsPart);

        /// <summary>
        /// Gets the image content type, such as image/png or image/jpeg.
        /// </summary>
        public string ContentType => ImagePart?.ContentType ?? string.Empty;

        /// <summary>
        /// Returns a copy of the image bytes from the worksheet drawing relationship.
        /// </summary>
        public byte[] GetBytes() {
            ImagePart? imagePart = ImagePart;
            if (imagePart == null) {
                return Array.Empty<byte>();
            }

            using Stream source = imagePart.GetStream();
            using var destination = new MemoryStream();
            source.CopyTo(destination);
            return destination.ToArray();
        }

        /// <summary>
        /// Sets image title and description metadata.
        /// </summary>
        public ExcelImage SetAltText(string description, string? title = null) {
            if (DrawingProperties != null) {
                DrawingProperties.Description = description ?? string.Empty;
                if (title != null) {
                    DrawingProperties.Title = title;
                }
                Save();
            }

            return this;
        }

        /// <summary>
        /// Clears image alt text metadata for decorative images.
        /// </summary>
        public ExcelImage Decorative() {
            if (DrawingProperties != null) {
                DrawingProperties.Description = string.Empty;
                DrawingProperties.Title = string.Empty;
                Save();
            }

            return this;
        }

        /// <summary>
        /// Controls whether Excel should keep the picture aspect ratio locked.
        /// </summary>
        public ExcelImage LockAspectRatio(bool locked = true) {
            var drawingProps = _picture.NonVisualPictureProperties?.NonVisualPictureDrawingProperties;
            if (drawingProps == null) {
                return this;
            }

            var locks = drawingProps.GetFirstChild<A.PictureLocks>();
            if (locks == null) {
                locks = new A.PictureLocks();
                drawingProps.Append(locks);
            }

            locks.NoChangeAspect = locked;
            Save();
            return this;
        }

        /// <summary>
        /// Sets the image size in pixels.
        /// </summary>
        public ExcelImage SetSize(int widthPixels, int heightPixels) {
            if (widthPixels <= 0) throw new ArgumentOutOfRangeException(nameof(widthPixels));
            if (heightPixels <= 0) throw new ArgumentOutOfRangeException(nameof(heightPixels));

            long cx = PxToEmu(widthPixels);
            long cy = PxToEmu(heightPixels);

            var extent = _anchor.GetFirstChild<Xdr.Extent>();
            if (extent != null) {
                extent.Cx = cx;
                extent.Cy = cy;
            }

            var transform = _picture.ShapeProperties?.GetFirstChild<A.Transform2D>();
            var transformExtents = transform?.GetFirstChild<A.Extents>();
            if (transformExtents != null) {
                transformExtents.Cx = cx;
                transformExtents.Cy = cy;
            }

            Save();
            return this;
        }

        private Xdr.NonVisualDrawingProperties? DrawingProperties
            => _picture.NonVisualPictureProperties?.NonVisualDrawingProperties;

        private A.SourceRectangle? CropRectangle
            => _picture.BlipFill?.SourceRectangle;

        private A.Transform2D? Transform
            => _picture.ShapeProperties?.GetFirstChild<A.Transform2D>();

        private ImagePart? ImagePart {
            get {
                string? relationshipId = _picture.BlipFill?.Blip?.Embed?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId)) {
                    return null;
                }

                try {
                    return _drawingsPart.GetPartById(relationshipId!) as ImagePart;
                } catch (ArgumentOutOfRangeException) {
                    return null;
                }
            }
        }

        private void Save() {
            _drawingsPart.WorksheetDrawing?.Save();
        }

        private static int GetDrawingOrder(OpenXmlElement anchor, DrawingsPart drawingsPart) {
            Xdr.WorksheetDrawing? worksheetDrawing = drawingsPart.WorksheetDrawing;
            if (worksheetDrawing == null) {
                return 0;
            }

            OpenXmlElementList children = worksheetDrawing.ChildElements;
            for (int i = 0; i < children.Count; i++) {
                if (ReferenceEquals(children[i], anchor)) {
                    return i;
                }
            }

            return 0;
        }

        private static long PxToEmu(int px) => (long)Math.Round(px * 9525.0);

        private static int EmuToPx(long emu) {
            if (emu <= 0) {
                return 0;
            }

            return (int)Math.Max(1, Math.Round(emu / 9525.0));
        }

        private static int EmuOffsetToPx(long emu) => (int)Math.Round(emu / 9525.0);

        private static double GetCropRatio(int? percentage) {
            if (!percentage.HasValue || percentage.Value <= 0) {
                return 0D;
            }

            return Math.Min(0.999D, percentage.Value / 100000D);
        }

        private static double GetRotationDegrees(int? angle) {
            if (!angle.HasValue || angle.Value == 0) {
                return 0D;
            }

            double degrees = angle.Value / 60000D;
            degrees %= 360D;
            if (degrees < 0D) {
                degrees += 360D;
            }

            return degrees;
        }

        private int GetMarkerRow() {
            string? row = _anchor.GetFirstChild<Xdr.FromMarker>()?.RowId?.Text;
            return int.TryParse(row, out int value) && value >= 0 ? value : 0;
        }

        private int GetMarkerColumn() {
            string? column = _anchor.GetFirstChild<Xdr.FromMarker>()?.ColumnId?.Text;
            return int.TryParse(column, out int value) && value >= 0 ? value : 0;
        }

        private long GetMarkerColumnOffset() {
            string? offset = _anchor.GetFirstChild<Xdr.FromMarker>()?.ColumnOffset?.Text;
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private long GetMarkerRowOffset() {
            string? offset = _anchor.GetFirstChild<Xdr.FromMarker>()?.RowOffset?.Text;
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private bool TryGetToMarkerRow(out int row) {
            string? value = _anchor.GetFirstChild<Xdr.ToMarker>()?.RowId?.Text;
            return int.TryParse(value, out row) && row >= 0;
        }

        private bool TryGetToMarkerColumn(out int column) {
            string? value = _anchor.GetFirstChild<Xdr.ToMarker>()?.ColumnId?.Text;
            return int.TryParse(value, out column) && column >= 0;
        }

        private long GetToMarkerColumnOffset() {
            string? offset = _anchor.GetFirstChild<Xdr.ToMarker>()?.ColumnOffset?.Text;
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private long GetToMarkerRowOffset() {
            string? offset = _anchor.GetFirstChild<Xdr.ToMarker>()?.RowOffset?.Text;
            return long.TryParse(offset, out long value) ? value : 0L;
        }

        private long GetExtentCx() {
            long? anchorExtent = _anchor.GetFirstChild<Xdr.Extent>()?.Cx?.Value;
            if (anchorExtent.HasValue && anchorExtent.Value > 0) {
                return anchorExtent.Value;
            }

            long? shapeExtent = _picture.ShapeProperties?.GetFirstChild<A.Transform2D>()?.GetFirstChild<A.Extents>()?.Cx?.Value;
            return shapeExtent.GetValueOrDefault();
        }

        private long GetExtentCy() {
            long? anchorExtent = _anchor.GetFirstChild<Xdr.Extent>()?.Cy?.Value;
            if (anchorExtent.HasValue && anchorExtent.Value > 0) {
                return anchorExtent.Value;
            }

            long? shapeExtent = _picture.ShapeProperties?.GetFirstChild<A.Transform2D>()?.GetFirstChild<A.Extents>()?.Cy?.Value;
            return shapeExtent.GetValueOrDefault();
        }
    }
}
