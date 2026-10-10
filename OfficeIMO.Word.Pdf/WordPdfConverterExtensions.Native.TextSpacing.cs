using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeTextSpacing(double? WidthPercentage, double? CharacterSpacing, double? KerningMinimumFontSize = null) {
            public NativeTextSpacing Merge(OpenXmlElement? properties) => Merge(new NativeTextSpacing(
                ReadNativeCharacterScale(properties?.GetFirstChild<W.CharacterScale>()),
                properties?.GetFirstChild<W.Spacing>()?.Val?.Value / 20D,
                properties?.GetFirstChild<W.Kern>()?.Val?.Value / 2D));

            public NativeTextSpacing Merge(NativeTextSpacing overrides) => new(
                overrides.WidthPercentage ?? WidthPercentage,
                overrides.CharacterSpacing ?? CharacterSpacing,
                overrides.KerningMinimumFontSize ?? KerningMinimumFontSize);

            public NativeTextSpacing Resolve() => new(WidthPercentage ?? 100D, CharacterSpacing ?? 0D, KerningMinimumFontSize);

            // Word uses half-points for the minimum size; zero explicitly disables
            // kerning. Missing formatting inherits, and an entirely absent setting
            // leaves Word kerning disabled, independently of the PDF shaper default.
            public OfficeTextFeatureSettings ResolveFeatureSettings(double fontSize) =>
                OfficeTextFeatureSettings.Default.With("kern",
                    KerningMinimumFontSize is > 0D && fontSize >= KerningMinimumFontSize.Value ? 1 : 0);

            public PdfCore.PdfTextRun ApplyTo(PdfCore.PdfTextRun run, double defaultFontSize) {
                double width = WidthPercentage ?? 100D;
                double spacing = CharacterSpacing ?? 0D;
                if (run.HorizontalTextScaling != width) run = run.WithHorizontalTextScaling(width);
                if (run.CharacterSpacing != spacing) run = run.WithCharacterSpacing(spacing);
                return run.WithFeatureSettings(ResolveFeatureSettings(run.FontSize ?? defaultFontSize));
            }
        }

        private static double? ReadNativeCharacterScale(W.CharacterScale? scale) {
            if (scale?.Val == null) return null;
            long value = scale.Val.Value;
            if (value < 1 || value > 600)
                throw new InvalidOperationException("Word character width must be between 1 and 600 percent.");
            return value;
        }

        private static PdfCore.PdfTextRun CopyNativeTextSpacing(PdfCore.PdfTextRun source, PdfCore.PdfTextRun target) =>
            target.WithHorizontalTextScaling(source.HorizontalTextScaling)
                .WithCharacterSpacing(source.CharacterSpacing)
                .WithFeatureSettings(source.FeatureSettings);

        private static NativeTextSpacing ResolveNativeListMarkerTextSpacing(WordDocumentTraversal.ListInfo info, NativeTextSpacing inherited) {
            if (info.MarkerCharacterScale is < 1 or > 600)
                throw new InvalidOperationException("Word character width must be between 1 and 600 percent.");
            // Native Word leaves numbering labels unkerned, including labels
            // whose numbering level or paragraph mark explicitly enables kern.
            return inherited.Merge(new NativeTextSpacing(info.MarkerCharacterScale, info.MarkerCharacterSpacingTwips / 20D,
                0D)).Resolve();
        }
    }
}
