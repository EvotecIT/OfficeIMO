namespace OfficeIMO.Excel.LegacyXls.Model {
    internal static class LegacyXlsSheetVisibilityDecoder {
        internal static LegacyXlsSheetVisibility? ToKind(byte visibility) {
            switch (visibility) {
                case 0x00: return LegacyXlsSheetVisibility.Visible;
                case 0x01: return LegacyXlsSheetVisibility.Hidden;
                case 0x02: return LegacyXlsSheetVisibility.VeryHidden;
                default: return null;
            }
        }

        internal static string ToName(byte visibility) {
            LegacyXlsSheetVisibility? kind = ToKind(visibility);
            return kind.HasValue ? kind.Value.ToString() : $"Visibility:0x{visibility:X2}";
        }
    }
}
