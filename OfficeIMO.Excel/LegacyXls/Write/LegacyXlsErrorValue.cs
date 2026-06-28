namespace OfficeIMO.Excel.LegacyXls.Write {
    /// <summary>
    /// Maps Open XML error text to BIFF error codes used by BoolErr and FormulaValue records.
    /// </summary>
    internal static class LegacyXlsErrorValue {
        internal static bool TryGetCode(string text, out byte errorCode) {
            switch ((text ?? string.Empty).Trim().ToUpperInvariant()) {
                case "#NULL!":
                    errorCode = 0x00;
                    return true;
                case "#DIV/0!":
                    errorCode = 0x07;
                    return true;
                case "#VALUE!":
                    errorCode = 0x0f;
                    return true;
                case "#REF!":
                    errorCode = 0x17;
                    return true;
                case "#NAME?":
                    errorCode = 0x1d;
                    return true;
                case "#NUM!":
                    errorCode = 0x24;
                    return true;
                case "#N/A":
                    errorCode = 0x2a;
                    return true;
                case "#GETTING_DATA":
                    errorCode = 0x2b;
                    return true;
                default:
                    errorCode = 0;
                    return false;
            }
        }
    }
}
