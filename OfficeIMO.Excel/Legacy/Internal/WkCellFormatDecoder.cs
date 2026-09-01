namespace OfficeIMO.Excel.Legacy;

internal static class WkCellFormatDecoder {
    internal static string? Decode(byte format, bool isText) {
        if (isText) return null;
        int style = (format >> 4) & 0x07;
        int digits = format & 0x0F;
        string decimals = digits == 0 ? "0" : "0." + new string('0', digits);
        switch (style) {
            case 0: return decimals;
            case 1: return decimals + "E+00";
            case 2: return "$#,##" + (digits == 0 ? "0" : "0." + new string('0', digits));
            case 3: return decimals + "%";
            case 4: return "#,##" + (digits == 0 ? "0" : "0." + new string('0', digits));
            case 5 when digits == 4: return "mm/dd/yy";
            case 5 when digits == 5: return "mm/dd";
            default: return null;
        }
    }
}
