namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Describes the kind of background fill currently assigned to a PowerPoint slide.
    /// </summary>
    public enum PowerPointSlideBackgroundKind {
        /// <summary>No explicit slide background fill is present.</summary>
        None,

        /// <summary>The slide background uses a solid RGB color.</summary>
        SolidColor,

        /// <summary>The slide background uses an embedded image.</summary>
        Image,

        /// <summary>The slide background uses a two-stop linear RGB gradient.</summary>
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
            string? gradientStartColor,
            string? gradientEndColor,
            double? gradientAngleDegrees,
            string? unsupportedReason) {
            Kind = kind;
            Color = color;
            _imageBytes = imageBytes == null ? null : (byte[])imageBytes.Clone();
            ImageContentType = imageContentType;
            GradientStartColor = gradientStartColor;
            GradientEndColor = gradientEndColor;
            GradientAngleDegrees = gradientAngleDegrees;
            UnsupportedReason = unsupportedReason;
        }

        /// <summary>Background kind.</summary>
        public PowerPointSlideBackgroundKind Kind { get; }

        /// <summary>Solid background color as a six-digit RGB hex value when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.SolidColor"/>.</summary>
        public string? Color { get; }

        /// <summary>Embedded background image bytes when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.Image"/>.</summary>
        public byte[]? ImageBytes => _imageBytes == null ? null : (byte[])_imageBytes.Clone();

        /// <summary>Embedded image content type when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.Image"/>.</summary>
        public string? ImageContentType { get; }

        /// <summary>First linear gradient color as a six-digit RGB hex value when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.LinearGradient"/>.</summary>
        public string? GradientStartColor { get; }

        /// <summary>Last linear gradient color as a six-digit RGB hex value when <see cref="Kind"/> is <see cref="PowerPointSlideBackgroundKind.LinearGradient"/>.</summary>
        public string? GradientEndColor { get; }

        /// <summary>PowerPoint linear gradient angle in degrees when available.</summary>
        public double? GradientAngleDegrees { get; }

        /// <summary>Explanation for unsupported background fills.</summary>
        public string? UnsupportedReason { get; }

        internal static PowerPointSlideBackground None() =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.None, null, null, null, null, null, null, null);

        internal static PowerPointSlideBackground SolidColor(string color) =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.SolidColor, color, null, null, null, null, null, null);

        internal static PowerPointSlideBackground Image(byte[] imageBytes, string? contentType) =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.Image, null, imageBytes, contentType, null, null, null, null);

        internal static PowerPointSlideBackground LinearGradient(string startColor, string endColor, double angleDegrees) =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.LinearGradient, null, null, null, startColor, endColor, angleDegrees, null);

        internal static PowerPointSlideBackground Unsupported(string reason) =>
            new PowerPointSlideBackground(PowerPointSlideBackgroundKind.Unsupported, null, null, null, null, null, null, reason);
    }
}
