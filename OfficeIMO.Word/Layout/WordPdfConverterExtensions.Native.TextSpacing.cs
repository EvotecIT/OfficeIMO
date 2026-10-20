using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeTextSpacing(double? WidthPercentage, double? CharacterSpacing,
            double? KerningMinimumFontSize = null, string? LigatureSetting = null) {
            public NativeTextSpacing Merge(OpenXmlElement? properties) => Merge(new NativeTextSpacing(
                ReadNativeCharacterScale(properties?.GetFirstChild<W.CharacterScale>()),
                properties?.GetFirstChild<W.Spacing>()?.Val?.Value / 20D,
                properties?.GetFirstChild<W.Kern>()?.Val?.Value / 2D,
                ReadNativeLigatureSetting(properties)));

            public NativeTextSpacing Merge(NativeTextSpacing overrides) => new(
                overrides.WidthPercentage ?? WidthPercentage,
                overrides.CharacterSpacing ?? CharacterSpacing,
                overrides.KerningMinimumFontSize ?? KerningMinimumFontSize,
                overrides.LigatureSetting ?? LigatureSetting);

            public NativeTextSpacing Resolve(bool openTypeFeaturesEnabled = true) => new(WidthPercentage ?? 100D, CharacterSpacing ?? 0D,
                KerningMinimumFontSize, openTypeFeaturesEnabled ? LigatureSetting : LigatureSetting == null ? null : "none");

            // Word uses half-points for the minimum size; zero explicitly disables
            // kerning. Missing formatting inherits, and an entirely absent setting
            // leaves Word kerning disabled, independently of the PDF shaper default.
            public OfficeTextFeatureSettings ResolveFeatureSettings(double fontSize) {
                string? setting = LigatureSetting;
                bool all = setting == "all";
                bool Enabled(string name) => all || setting?.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0;
                return OfficeTextFeatureSettings.Default.With("kern",
                        KerningMinimumFontSize is > 0D && fontSize >= KerningMinimumFontSize.Value ? 1 : 0)
                    .With("liga", Enabled("standard") ? 1 : 0)
                    .With("clig", Enabled("contextual") ? 1 : 0)
                    .With("dlig", Enabled("discretional") ? 1 : 0)
                    .With("hlig", Enabled("historical") ? 1 : 0);
            }

            public PdfCore.PdfTextRun ApplyTo(PdfCore.PdfTextRun run, double defaultFontSize) {
                double width = WidthPercentage ?? 100D;
                double spacing = CharacterSpacing ?? 0D;
                if (run.HorizontalTextScaling != width) run = run.WithHorizontalTextScaling(width);
                if (run.CharacterSpacing != spacing) run = run.WithCharacterSpacing(spacing);
                return run.WithFeatureSettings(ResolveFeatureSettings(run.FontSize ?? defaultFontSize));
            }
        }

        private static string? ReadNativeLigatureSetting(OpenXmlElement? properties) {
            const string word2010 = "http://schemas.microsoft.com/office/word/2010/wordml";
            // The SDK retains this extension as an unknown element in some
            // run-property containers, including loaded document defaults.
            OpenXmlElement? ligatures = properties?.ChildElements
                .LastOrDefault(child => child.LocalName == "ligatures" && child.NamespaceUri == word2010);
            return ligatures?.GetAttributes()
                .FirstOrDefault(attribute => attribute.LocalName == "val" && attribute.NamespaceUri == word2010).Value;
        }

        private static double? ReadNativeCharacterScale(W.CharacterScale? scale) {
            if (scale?.Val == null) return null;
            long value = scale.Val.Value;
            if (value < 1 || value > 600)
                throw new InvalidOperationException("Word character width must be between 1 and 600 percent.");
            return value;
        }

        private static PdfCore.PdfTextRun CopyNativeTextSpacing(PdfCore.PdfTextRun source, PdfCore.PdfTextRun target) =>
            target.WithSpacingFrom(source);

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
