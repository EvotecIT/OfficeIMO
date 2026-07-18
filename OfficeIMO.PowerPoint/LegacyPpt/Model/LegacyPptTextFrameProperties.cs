using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>
    /// Represents the classic OfficeArt text-frame properties attached to a
    /// text-bearing shape.
    /// </summary>
    public sealed class LegacyPptTextFrameProperties {
        private LegacyPptTextFrameProperties(
            IReadOnlyList<OfficeArtProperty> properties) {
            LeftInsetEmus = GetCoordinate(properties, 0x0081);
            TopInsetEmus = GetCoordinate(properties, 0x0082);
            RightInsetEmus = GetCoordinate(properties, 0x0083);
            BottomInsetEmus = GetCoordinate(properties, 0x0084);
            WrapMode = GetUInt32(properties, 0x0085);
            AnchorMode = GetUInt32(properties, 0x0087);
            TextFlow = GetUInt32(properties, 0x0088);
            uint? textFlags = GetUInt32(properties, 0x00BF);
            AutoTextMargin = GetBoolean(textFlags, 1U << 12, 1U << 28);
            FitShapeToText = GetBoolean(textFlags, 1U << 14, 1U << 30);
            CanRewriteProjectedProperties =
                (!WrapMode.HasValue || WrapMode.Value is 0 or 2)
                && (!AnchorMode.HasValue || AnchorMode.Value <= 5)
                && (!TextFlow.HasValue || TextFlow.Value <= 2
                    || TextFlow.Value == 4);
        }

        internal static LegacyPptTextFrameProperties Decode(
            IReadOnlyList<OfficeArtProperty> properties) => new(
                properties ?? Array.Empty<OfficeArtProperty>());

        /// <summary>Gets the explicit left text inset in EMUs.</summary>
        public int? LeftInsetEmus { get; }

        /// <summary>Gets the explicit top text inset in EMUs.</summary>
        public int? TopInsetEmus { get; }

        /// <summary>Gets the explicit right text inset in EMUs.</summary>
        public int? RightInsetEmus { get; }

        /// <summary>Gets the explicit bottom text inset in EMUs.</summary>
        public int? BottomInsetEmus { get; }

        /// <summary>Gets the raw MSOWRAPMODE value.</summary>
        public uint? WrapMode { get; }

        /// <summary>Gets the raw MSOANCHOR value.</summary>
        public uint? AnchorMode { get; }

        /// <summary>Gets the raw MSOTXFL text-flow value.</summary>
        public uint? TextFlow { get; }

        /// <summary>
        /// Gets whether OfficeArt explicitly requests automatic text margins.
        /// </summary>
        public bool? AutoTextMargin { get; }

        /// <summary>
        /// Gets whether OfficeArt explicitly requests resizing the shape to
        /// fit its text.
        /// </summary>
        public bool? FitShapeToText { get; }

        /// <summary>
        /// Gets whether every projected text-frame value can be written back
        /// without collapsing a distinct classic enum value.
        /// </summary>
        public bool CanRewriteProjectedProperties { get; }

        /// <summary>Gets whether the shape carries explicit frame state.</summary>
        public bool HasExplicitProperties => LeftInsetEmus.HasValue
            || TopInsetEmus.HasValue || RightInsetEmus.HasValue
            || BottomInsetEmus.HasValue || WrapMode.HasValue
            || AnchorMode.HasValue || TextFlow.HasValue
            || AutoTextMargin.HasValue || FitShapeToText.HasValue;

        private static int? GetCoordinate(
            IReadOnlyList<OfficeArtProperty> properties,
            ushort propertyId) {
            uint? value = GetUInt32(properties, propertyId);
            if (!value.HasValue) return null;
            int coordinate = unchecked((int)value.Value);
            return coordinate >= 0 ? coordinate : null;
        }

        private static uint? GetUInt32(
            IReadOnlyList<OfficeArtProperty> properties,
            ushort propertyId) => properties.LastOrDefault(property =>
                property.PropertyId == propertyId && !property.IsComplex)?
                .Value;

        private static bool? GetBoolean(uint? value, uint useMask,
            uint valueMask) => !value.HasValue
                || (value.Value & useMask) == 0
                    ? null
                    : (value.Value & valueMask) != 0;
    }
}
