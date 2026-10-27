using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeComplexScriptDefaults(double? FontSize, bool? Enabled) {
            public NativeComplexScriptDefaults Merge(OpenXmlElement? properties) => new(
                ParseNativeDefaultFontSize(properties?.GetFirstChild<W.FontSizeComplexScript>()?.Val?.Value) ?? FontSize,
                ReadNativeOnOff(properties?.GetFirstChild<W.ComplexScript>()) ??
                    ReadNativeOnOff(properties?.GetFirstChild<W.RightToLeftText>()) ?? Enabled);

            public NativeComplexScriptDefaults Merge(NativeComplexScriptDefaults overrides) => new(
                overrides.FontSize ?? FontSize, overrides.Enabled ?? Enabled);
        }

        private static double? ResolveNativeComplexScriptFontSize(
            double? ordinarySize,
            W.RunProperties? properties,
            NativeCharacterStyleDefaults characterStyle,
            NativeParagraphStyleDefaults paragraphStyle,
            NativeTableRunStyleDefaults tableStyle,
            NativeDocumentDefaults documentDefaults) {
            NativeComplexScriptDefaults resolved = documentDefaults.ComplexScript
                .Merge(tableStyle.ComplexScript)
                .Merge(paragraphStyle.ComplexScript)
                .Merge(characterStyle.ComplexScript)
                .Merge(properties);
            return resolved.Enabled == true ? resolved.FontSize ?? ordinarySize : ordinarySize;
        }
    }
}
