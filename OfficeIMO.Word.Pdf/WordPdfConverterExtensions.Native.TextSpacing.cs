using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeTextSpacing(double? WidthPercentage, double? CharacterSpacing) {
            public NativeTextSpacing Merge(OpenXmlElement? properties) => Merge(new NativeTextSpacing(
                ReadNativeCharacterScale(properties?.GetFirstChild<W.CharacterScale>()),
                properties?.GetFirstChild<W.Spacing>()?.Val?.Value / 20D));

            public NativeTextSpacing Merge(NativeTextSpacing overrides) => new(
                overrides.WidthPercentage ?? WidthPercentage,
                overrides.CharacterSpacing ?? CharacterSpacing);

            public NativeTextSpacing Resolve() => new(WidthPercentage ?? 100D, CharacterSpacing ?? 0D);

            public PdfCore.PdfTextRun ApplyTo(PdfCore.PdfTextRun run) {
                double width = WidthPercentage ?? 100D;
                double spacing = CharacterSpacing ?? 0D;
                if (run.HorizontalTextScaling != width) run = run.WithHorizontalTextScaling(width);
                if (run.CharacterSpacing != spacing) run = run.WithCharacterSpacing(spacing);
                return run;
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
            new NativeTextSpacing(source.HorizontalTextScaling, source.CharacterSpacing).ApplyTo(target);
    }
}
