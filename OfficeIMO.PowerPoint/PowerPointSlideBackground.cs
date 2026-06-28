namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Describes the kind of background fill currently assigned to a PowerPoint slide.
    /// </summary>
    public enum PowerPointSlideBackgroundKind {
        /// <summary>No explicit slide background fill is present.</summary>
        None,

        /// <summary>The slide background uses a solid RGB or RGBA color.</summary>
        SolidColor,

        /// <summary>The slide background uses an embedded image.</summary>
        Image,

        /// <summary>The slide background uses a two-stop linear RGB or RGBA gradient.</summary>
        LinearGradient,

        /// <summary>The slide background uses an Office fill that is not yet exposed as reusable data.</summary>
        Unsupported
    }

    /// <summary>
    /// Immutable snapshot of a PowerPoint slide background for exporters such as PDF.
    /// </summary>
    public sealed class PowerPointSlideBackground {
        private readonly byte[]? _imageBytes;

        private PowerPointSlideBackground(
            PowerPointSlideBackgroundKind kind,
            string? color,
            byte[]? imageBytes,
            string? imageContentType,
            PowerPointPictureCrop imageCrop,
            string? gradientStartColor,
            string? gradientEndColor,
            double? gradientAngleDegrees,
            string? unsupportedReason) {
            Kind = kind;
            Color = color;
            _imageBytes = imageBytes == null ? null : (byte[])imageBytes.Clone();
            ImageContentType = imageContentType;
            ImageCropLeft = imageCrop.Left;
            ImageCropTop = imageCrop.Top;
            ImageCropRight = imageCrop.Right;
            ImageCropBottom = imageCrop.Bottom;
            GradientStartColor = gradientStartColor;
            GradientEndColor = gradientEndColor;
            GradientAngleDegrees = gradientAngleDegrees;
            UnsupportedReason = unsupportedReason;
        }

        /// <summary>Background kind.</summary>
        public PowerPointSlideBackgroundKind Kind { get; }

        /// <summary>Solid background color as RGB or RGBA hex when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.SolidColor"/>.</summary>
        public string? Color { get; }

        /// <summary>Embedded background image bytes when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.Image"/>.</summary>
        public byte[]? ImageBytes => _imageBytes == null ? null : (byte[])_imageBytes.Clone();

        /// <summary>Embedded image content type when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.Image"/>.</summary>
        public string? ImageContentType { get; }

        /// <summary>Fraction cropped from the embedded background image left edge.</summary>
        public double ImageCropLeft { get; }

        /// <summary>Fraction cropped from the embedded background image top edge.</summary>
        public double ImageCropTop { get; }

        /// <summary>Fraction cropped from the embedded background image right edge.</summary>
        public double ImageCropRight { get; }

        /// <summary>Fraction cropped from the embedded background image bottom edge.</summary>
        public double ImageCropBottom { get; }

        /// <summary>True when the image background carries a source crop.</summary>
        public bool HasImageCrop => ImageCropLeft > 0D || ImageCropTop > 0D || ImageCropRight > 0D || ImageCropBottom > 0D;

        /// <summary>First linear gradient color as RGB or RGBA hex when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.LinearGradient"/>.</summary>
        public string? GradientStartColor { get; }

        /// <summary>Last linear gradient color as RGB or RGBA hex when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.LinearGradient"/>.</summary>
        public string? GradientEndColor { get; }

        /// <summary>PowerPoint linear gradient angle in degrees when available.</summary>
        public double? GradientAngleDegrees { get; }

        /// <summary>Explanation for unsupported background fills.</summary>
        public string? UnsupportedReason { get; }

        internal static PowerPointSlideBackground None() =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.None, null, null, null, PowerPointPictureCrop.None, null, null, null, null);

        internal static PowerPointSlideBackground SolidColor(string color) =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.SolidColor, color, null, null, PowerPointPictureCrop.None, null, null, null, null);

        internal static PowerPointSlideBackground Image(byte[] imageBytes, string? contentType, PowerPointPictureCrop imageCrop = default) =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.Image, null, imageBytes, contentType, imageCrop.Equals(default(PowerPointPictureCrop)) ? PowerPointPictureCrop.None : imageCrop, null, null, null, null);

        internal static PowerPointSlideBackground LinearGradient(string startColor, string endColor, double angleDegrees) =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.LinearGradient, null, null, null, PowerPointPictureCrop.None, startColor, endColor, angleDegrees, null);

        internal static PowerPointSlideBackground Unsupported(string reason) =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.Unsupported, null, null, null, PowerPointPictureCrop.None, null, null, null, reason);
    }
}
