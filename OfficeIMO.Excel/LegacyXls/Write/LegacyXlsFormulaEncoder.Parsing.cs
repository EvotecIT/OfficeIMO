namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsFormulaEncoder {
        private static bool TrySkipQuotedSheetName(string text, ref int index) {
            if (index < 0 || index >= text.Length || text[index] != '\'') {
                return false;
            }

            for (int i = index + 1; i < text.Length; i++) {
                if (text[i] != '\'') {
                    continue;
                }

                if (i + 1 < text.Length && text[i + 1] == '\'') {
                    i++;
                    continue;
                }

                index = i;
                return true;
            }

            return false;
        }

        private static bool TrySkipQuotedSheetNameBackward(string text, ref int index) {
            if (index < 0 || index >= text.Length || text[index] != '\'') {
                return false;
            }

            for (int i = index - 1; i >= 0; i--) {
                if (text[i] != '\'') {
                    continue;
                }

                if (i > 0 && text[i - 1] == '\'') {
                    i--;
                    continue;
                }

                index = i;
                return true;
            }

            return false;
        }
    }
}
