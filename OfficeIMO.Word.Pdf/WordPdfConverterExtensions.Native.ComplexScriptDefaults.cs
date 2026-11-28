namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private readonly record struct NativeComplexScriptDefaults(double? FontSize, bool? Enabled) {
            public NativeComplexScriptDefaults Merge(OfficeXmlNode properties) => new(
                ParseNativeDefaultFontSize(properties.WChild("szCs").WVal()) ?? FontSize,
                ReadNativeOnOff(properties.WChild("cs")) ??
                    ReadNativeOnOff(properties.WChild("rtl")) ?? Enabled);

            public NativeComplexScriptDefaults Merge(NativeComplexScriptDefaults overrides) => new(
                overrides.FontSize ?? FontSize, overrides.Enabled ?? Enabled);
        }

        private static double? ResolveNativeComplexScriptFontSize(
            double? ordinarySize,
            OfficeXmlNode properties,
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
