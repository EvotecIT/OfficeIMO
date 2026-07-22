namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    /// <summary>
    /// Central resource ceilings for binary PowerPoint conversion paths that
    /// materialize caller-controlled package content.
    /// </summary>
    internal static partial class LegacyPptWriter {
        internal const int MaximumGroupNestingDepth = 64;
        internal const long MaximumStaticVisualRasterPixels = 16_000_000L;
        internal const int MaximumPictureBytes = 64 * 1024 * 1024;
        internal const long MaximumTotalPictureBytes = 256L * 1024 * 1024;
        internal const int MaximumSoundBytes = 64 * 1024 * 1024;
        internal const long MaximumTotalSoundBytes = 256L * 1024 * 1024;
    }
}
