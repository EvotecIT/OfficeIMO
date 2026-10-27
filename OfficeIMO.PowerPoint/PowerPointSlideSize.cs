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
            Presentation presentation = _presentationPart.Presentation ??= new Presentation();
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
            get => OfficeLength.Emus(EnsureSlideSize().Cx?.Value ?? 0);
            set => EnsureSlideSize().Cx = EnsureEmuFitsInt(value.ToEmus(), nameof(Width));
        }

        /// <summary>Slide height.</summary>
        public OfficeLength Height {
            get => OfficeLength.Emus(EnsureSlideSize().Cy?.Value ?? 0);
            set => EnsureSlideSize().Cy = EnsureEmuFitsInt(value.ToEmus(), nameof(Height));
        }

        /// <summary>Gets the slide width divided by its height.</summary>
        public double AspectRatio => Height == OfficeLength.Zero ? 0 : Width.ToEmus() / (double)Height.ToEmus();
        /// <summary>Gets whether the slide is in portrait orientation.</summary>
        public bool IsPortrait => Height > Width;
        /// <summary>Gets whether the slide is in landscape orientation.</summary>
        public bool IsLandscape => Width >= Height;

        /// <summary>Slide size preset type.</summary>
        public PowerPointSlideSizeType? Type {
            get => EnsureSlideSize().Type?.Value.ToOfficeEnum();
            set => EnsureSlideSize().Type = value?.ToOpenXml();
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
            if (columnCount <= 0) throw new ArgumentOutOfRangeException(nameof(columnCount));
            return GetContentBox(margin).SplitColumns(columnCount, gutter);
        }

        /// <summary>Creates equal rows inside the slide's content region.</summary>
        public PowerPointLayoutBox[] GetRows(int rowCount, OfficeLength margin, OfficeLength gutter) {
            if (rowCount <= 0) throw new ArgumentOutOfRangeException(nameof(rowCount));
            return GetContentBox(margin).SplitRows(rowCount, gutter);
        }
    }
}
