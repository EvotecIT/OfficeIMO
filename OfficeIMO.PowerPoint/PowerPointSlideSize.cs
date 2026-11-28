using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Provides access to presentation slide size in various units.
    /// </summary>
    public sealed class PowerPointSlideSize {
        private readonly PowerPointDocument _document;

        internal PowerPointSlideSize(PowerPointDocument document) => _document = document;

        private PSlideSizeView GetSize(PowerPointPartRef part) {
            var presentation = new PPresentationView(part.Root);
            if (presentation.SldSz is { } existing) return existing;
            PSlideSizeView size = presentation.GetOrAddSldSz();
            size.Cx = 12192000;
            size.Cy = 6858000;
            size.Type = PowerPointSlideSizeType.Screen16x9.ToOpenXmlToken();
            part.MarkEdited();
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
            get => OfficeLength.Emus(GetSize(_document.PresentationPartRef).Cx ?? 0);
            set {
                int width = EnsureEmuFitsInt(value.ToEmus(), nameof(Width));
                PowerPointPartRef part = _document.PresentationPartRef;
                PSlideSizeView size = GetSize(part);
                size.Cx = width;
                part.MarkEdited();
            }
        }

        /// <summary>Slide height.</summary>
        public OfficeLength Height {
            get => OfficeLength.Emus(GetSize(_document.PresentationPartRef).Cy ?? 0);
            set {
                int height = EnsureEmuFitsInt(value.ToEmus(), nameof(Height));
                PowerPointPartRef part = _document.PresentationPartRef;
                PSlideSizeView size = GetSize(part);
                size.Cy = height;
                part.MarkEdited();
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
            get => GetSize(_document.PresentationPartRef).Type is { } token
                && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointSlideSizeType type) ? type : null;
            set {
                string? token = value?.ToOpenXmlToken();
                PowerPointPartRef part = _document.PresentationPartRef;
                PSlideSizeView size = GetSize(part);
                size.Type = token;
                part.MarkEdited();
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
