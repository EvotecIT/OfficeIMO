using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one resolved color stop in a legacy OfficeArt gradient.</summary>
    public sealed class LegacyPptGradientStop {
        internal LegacyPptGradientStop(string? color, double position) {
            Color = color;
            Position = position;
        }

        /// <summary>Gets the resolved color as RRGGBB, or null when its source reference is unresolved.</summary>
        public string? Color { get; }

        /// <summary>Gets the relative position from zero through one.</summary>
        public double Position { get; }
    }

    /// <summary>Identifies the OfficeArt fill used by a binary slide or master background.</summary>
    public enum LegacyPptBackgroundKind {
        /// <summary>The background shape explicitly disables its fill.</summary>
        None,
        /// <summary>A solid color.</summary>
        Solid,
        /// <summary>A foreground/background pattern.</summary>
        Pattern,
        /// <summary>A tiled texture image.</summary>
        Texture,
        /// <summary>A stretched picture.</summary>
        Picture,
        /// <summary>A linear gradient.</summary>
        LinearGradient,
        /// <summary>A centered path gradient.</summary>
        CenterGradient,
        /// <summary>A gradient following the shape geometry.</summary>
        ShapeGradient,
        /// <summary>A scaled gradient.</summary>
        ScaleGradient,
        /// <summary>A title-area gradient.</summary>
        TitleGradient,
        /// <summary>The OfficeArt background-fill default.</summary>
        Inherited,
        /// <summary>An undefined fill type retained by its raw value.</summary>
        Unsupported
    }

    /// <summary>Represents an OfficeArt background shape and its decoded fill.</summary>
    public sealed class LegacyPptBackground {
        internal LegacyPptBackground(LegacyPptBackgroundKind kind, uint rawFillType,
            string? foregroundColor, string? backgroundColor, double? foregroundOpacity,
            double? backgroundOpacity, double? angleDegrees, int? focusPercent,
            int? pictureStoreIndex, OfficeArtBlipStoreEntry? picture,
            IReadOnlyList<LegacyPptGradientStop>? gradientStops,
            bool isGradientStopTableTruncated) {
            Kind = kind;
            RawFillType = rawFillType;
            ForegroundColor = foregroundColor;
            BackgroundColor = backgroundColor;
            ForegroundOpacity = foregroundOpacity;
            BackgroundOpacity = backgroundOpacity;
            AngleDegrees = angleDegrees;
            FocusPercent = focusPercent;
            PictureStoreIndex = pictureStoreIndex;
            Picture = picture;
            GradientStops = gradientStops?.ToArray() ?? Array.Empty<LegacyPptGradientStop>();
            IsGradientStopTableTruncated = isGradientStopTableTruncated;
        }

        /// <summary>Gets the typed background fill kind.</summary>
        public LegacyPptBackgroundKind Kind { get; }
        /// <summary>Gets the raw MSOFILLTYPE value.</summary>
        public uint RawFillType { get; }
        /// <summary>Gets the resolved foreground or first gradient color as RRGGBB.</summary>
        public string? ForegroundColor { get; }
        /// <summary>Gets the resolved background or last gradient color as RRGGBB.</summary>
        public string? BackgroundColor { get; }
        /// <summary>Gets foreground opacity from zero through one.</summary>
        public double? ForegroundOpacity { get; }
        /// <summary>Gets background opacity from zero through one.</summary>
        public double? BackgroundOpacity { get; }
        /// <summary>Gets the signed counterclockwise OfficeArt gradient angle.</summary>
        public double? AngleDegrees { get; }
        /// <summary>Gets the gradient focus position from -100 through 100.</summary>
        public int? FocusPercent { get; }
        /// <summary>Gets the one-based BLIP store index for image-backed fills.</summary>
        public int? PictureStoreIndex { get; }
        /// <summary>Gets the resolved BLIP store entry for image-backed fills.</summary>
        public OfficeArtBlipStoreEntry? Picture { get; }
        /// <summary>Gets resolved custom OfficeArt gradient stops.</summary>
        public IReadOnlyList<LegacyPptGradientStop> GradientStops { get; }
        /// <summary>Gets whether the declared gradient-stop table is malformed or truncated.</summary>
        public bool IsGradientStopTableTruncated { get; }

        /// <summary>Gets whether the background can be projected without dropping its primary fill.</summary>
        public bool HasProjectableFill => Kind == LegacyPptBackgroundKind.None
            || Kind == LegacyPptBackgroundKind.Inherited
            || Kind == LegacyPptBackgroundKind.Solid && ForegroundColor != null
            || (Kind is LegacyPptBackgroundKind.LinearGradient
                or LegacyPptBackgroundKind.CenterGradient
                or LegacyPptBackgroundKind.ShapeGradient
                or LegacyPptBackgroundKind.ScaleGradient
                or LegacyPptBackgroundKind.TitleGradient)
                && ForegroundColor != null && BackgroundColor != null
            || (Kind is LegacyPptBackgroundKind.Texture or LegacyPptBackgroundKind.Picture)
                && Picture?.HasImportableImage == true;
    }
}
