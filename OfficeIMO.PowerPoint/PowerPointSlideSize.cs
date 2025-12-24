using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Provides access to presentation slide size in various units.
    /// </summary>
    public sealed class PowerPointSlideSize {
        private readonly PresentationPart _presentationPart;

        internal PowerPointSlideSize(PresentationPart presentationPart) {
            _presentationPart = presentationPart ?? throw new ArgumentNullException(nameof(presentationPart));
        }

        private SlideSize EnsureSlideSize() {
            _presentationPart.Presentation ??= new Presentation();
            SlideSize? size = _presentationPart.Presentation.SlideSize;
            if (size == null) {
                size = new SlideSize { Cx = 12192000, Cy = 6858000, Type = SlideSizeValues.Screen16x9 };
                _presentationPart.Presentation.SlideSize = size;
            }
            return size;
        }

        /// <summary>
        ///     Slide width in EMUs.
        /// </summary>
        public long WidthEmus {
            get => EnsureSlideSize().Cx?.Value ?? 0;
            set => EnsureSlideSize().Cx = checked((int)value);
        }

        /// <summary>
        ///     Slide height in EMUs.
        /// </summary>
        public long HeightEmus {
            get => EnsureSlideSize().Cy?.Value ?? 0;
            set => EnsureSlideSize().Cy = checked((int)value);
        }

        /// <summary>
        ///     Slide width in centimeters.
        /// </summary>
        public double WidthCm {
            get => PowerPointUnits.ToCentimeters(WidthEmus);
            set => WidthEmus = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Slide height in centimeters.
        /// </summary>
        public double HeightCm {
            get => PowerPointUnits.ToCentimeters(HeightEmus);
            set => HeightEmus = PowerPointUnits.FromCentimeters(value);
        }

        /// <summary>
        ///     Slide width in inches.
        /// </summary>
        public double WidthInches {
            get => PowerPointUnits.ToInches(WidthEmus);
            set => WidthEmus = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Slide height in inches.
        /// </summary>
        public double HeightInches {
            get => PowerPointUnits.ToInches(HeightEmus);
            set => HeightEmus = PowerPointUnits.FromInches(value);
        }

        /// <summary>
        ///     Slide width in points.
        /// </summary>
        public double WidthPoints {
            get => PowerPointUnits.ToPoints(WidthEmus);
            set => WidthEmus = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Slide height in points.
        /// </summary>
        public double HeightPoints {
            get => PowerPointUnits.ToPoints(HeightEmus);
            set => HeightEmus = PowerPointUnits.FromPoints(value);
        }

        /// <summary>
        ///     Slide size preset type.
        /// </summary>
        public SlideSizeValues? Type {
            get => EnsureSlideSize().Type?.Value;
            set => EnsureSlideSize().Type = value;
        }
    }
}
