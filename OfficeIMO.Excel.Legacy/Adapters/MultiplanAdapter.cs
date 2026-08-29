namespace OfficeIMO.Excel.Legacy;

internal sealed class MultiplanAdapter : LegacySpreadsheetAdapterBase {
    public override LegacySpreadsheetFormat Format => LegacySpreadsheetFormat.Multiplan;
    public override string ProfileId => "microsoft-multiplan-dos-1-3";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        if (OfficeLegacyImportBuffer.StartsWith(data, 0x08, 0xE7) || OfficeLegacyImportBuffer.StartsWith(data, 0x0C, 0xEC) || OfficeLegacyImportBuffer.StartsWith(data, 0x0C, 0xED)) {
            reason = "Microsoft Multiplan DOS version signature.";
            return 100;
        }
        if (ExtensionIs(sourceName, ".mp", ".mp1", ".mp2", ".mp3")) {
            reason = "Microsoft Multiplan source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No Multiplan signature evidence.";
        return 0;
    }

    public override LegacySpreadsheetModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) =>
        ParseDelimitedSalvage(data, limits,
            "Multiplan text and tabular runs were salvaged; cell addresses, cached values, formulas, names, formatting, comments, and charts were not reconstructed.", cancellationToken);
}
