using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Provides access to presentation slide size in various units.
    /// </summary>
    public sealed class PowerPointSlideSize {
        private readonly PresentationPart? _sdkPresentationPart;
        private readonly PowerPointDocument? _document;

        internal PowerPointSlideSize(PresentationPart presentationPart) {
            _sdkPresentationPart = presentationPart ?? throw new ArgumentNullException(nameof(presentationPart));
        }

        // The slide size of a presentation that may not have its SDK package open: read from the package while it can be.
        internal PowerPointSlideSize(PowerPointDocument document) {
            _document = document;
        }

        private PresentationPart PresentationPart => _sdkPresentationPart ?? _document!.SdkPresentationPartForSize;

        // The slide size element's values in the package, when the presentation has one there.
        private bool TryReadPackage(out int width, out int height, out PowerPointSlideSizeType? type) {
            width = height = 0;
            type = null;
            if (_document?.GetPackageSlideSize() is not { } size) return false;
            if (!int.TryParse(size.Width ?? "0", System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out width)
                || !int.TryParse(size.Height ?? "0", System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out height)) {
                return false;
            }
            if (size.Type != null) {
                if (!PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(size.Type, out PowerPointSlideSizeType parsed)) return false;
                type = parsed;
            }
            return true;
        }

        private SlideSize EnsureSlideSize() {
            Presentation presentation = PresentationPart.Presentation ??= new Presentation();
            SlideSize? size = presentation.SlideSize;
            if (size == null) {
                size = new SlideSize { Cx = 12192000, Cy = 6858000, Type = SlideSizeValues.Screen16x9 };
                presentation.SlideSize = size;
            }
            return size;
        }

        private static int EnsureEmuFitsInt(long value, string paramName) {
            if (value < int.MinValue || value > int.MaxValue) {
                throw new ArgumentOutOfRangeException(paramName,
                    $"Value must be between {int.MinValue} and {int.MaxValue}.");
            }
            return (int)value;
        }
        /// <summary>Slide width.</summary>
        public OfficeLength Width {
            get => OfficeLength.Emus(TryReadPackage(out int width, out _, out _) ? width : EnsureSlideSize().Cx?.Value ?? 0);
            set {
                int width = EnsureEmuFitsInt(value.ToEmus(), nameof(Width));
                if (_document?.TrySetPackageSlideSize("cx", width.ToString(System.Globalization.CultureInfo.InvariantCulture)) == true) return;
                EnsureSlideSize().Cx = width;
            }
        }

        /// <summary>Slide height.</summary>
        public OfficeLength Height {
            get => OfficeLength.Emus(TryReadPackage(out _, out int height, out _) ? height : EnsureSlideSize().Cy?.Value ?? 0);
            set {
                int height = EnsureEmuFitsInt(value.ToEmus(), nameof(Height));
                if (_document?.TrySetPackageSlideSize("cy", height.ToString(System.Globalization.CultureInfo.InvariantCulture)) == true) return;
                EnsureSlideSize().Cy = height;
            }
        }

        /// <summary>Gets the slide width divided by its height.</summary>
        public double AspectRatio => Height == OfficeLength.Zero ? 0 : Width.ToEmus() / (double)Height.ToEmus();
        /// <summary>Gets whether the slide is in portrait orientation.</summary>
        public bool IsPortrait => Height > Width;
        /// <summary>Gets whether the slide is in landscape orientation.</summary>
        public bool IsLandscape => Width >= Height;

        /// <summary>Slide size preset type.</summary>
        public PowerPointSlideSizeType? Type {
            get => TryReadPackage(out _, out _, out PowerPointSlideSizeType? type) ? type : EnsureSlideSize().Type?.Value.ToOfficeEnum();
            set {
                if (_document?.TrySetPackageSlideSize("type", value.HasValue ? ((DocumentFormat.OpenXml.IEnumValue)value.Value.ToOpenXml()).Value : null) == true) return;
                EnsureSlideSize().Type = value?.ToOpenXml();
            }
        }

        /// <summary>Sets a slide size preset.</summary>
        public void SetPreset(PowerPointSlideSizePreset preset, bool portrait = false) {
            (OfficeLength width, OfficeLength height, PowerPointSlideSizeType type) = preset switch {
                PowerPointSlideSizePreset.Screen4x3 =>
                    (OfficeLength.Inches(10), OfficeLength.Inches(7.5), PowerPointSlideSizeType.Screen4x3),
                PowerPointSlideSizePreset.Screen16x9 =>
                    (OfficeLength.Inches(13.333), OfficeLength.Inches(7.5), PowerPointSlideSizeType.Screen16x9),
                PowerPointSlideSizePreset.Screen16x10 =>
                    (OfficeLength.Inches(10), OfficeLength.Inches(6.25), PowerPointSlideSizeType.Screen16x10),
                _ => throw new ArgumentOutOfRangeException(nameof(preset))
            };
            if (portrait) (width, height) = (height, width);
            SetSize(width, height, type);
        }

        /// <summary>Sets a custom slide size with explicit lengths.</summary>
        public void SetSize(OfficeLength width, OfficeLength height) => SetSize(width, height, PowerPointSlideSizeType.Custom);

        /// <summary>Sets the slide size and its preset type.</summary>
        public void SetSize(OfficeLength width, OfficeLength height, PowerPointSlideSizeType type) {
            if (width <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (height <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            EnsureEmuFitsInt(width.ToEmus(), nameof(width));
            EnsureEmuFitsInt(height.ToEmus(), nameof(height));
            Width = width;
            Height = height;
            Type = type;
        }

        /// <summary>Gets a content region that respects the specified margin.</summary>
        public PowerPointLayoutBox GetContentBox(OfficeLength margin) {
            if (margin < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(margin));
            if (margin.ToEmus() > Width.ToEmus() / 2 || margin.ToEmus() > Height.ToEmus() / 2)
                throw new ArgumentOutOfRangeException(nameof(margin), "Margin exceeds slide size.");
            return new PowerPointLayoutBox(margin, margin, Width - margin * 2, Height - margin * 2);
        }

        /// <summary>Creates equal columns inside the slide's content region.</summary>
        public PowerPointLayoutBox[] GetColumns(int columnCount, OfficeLength margin, OfficeLength gutter) {
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(columnCount);
            return GetContentBox(margin).SplitColumns(columnCount, gutter);
        }

        /// <summary>Creates equal rows inside the slide's content region.</summary>
        public PowerPointLayoutBox[] GetRows(int rowCount, OfficeLength margin, OfficeLength gutter) {
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(rowCount);
            return GetContentBox(margin).SplitRows(rowCount, gutter);
        }
    }
}
