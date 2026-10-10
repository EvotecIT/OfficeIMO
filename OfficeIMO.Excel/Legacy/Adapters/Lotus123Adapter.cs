namespace OfficeIMO.Excel.Legacy;

internal sealed class Lotus123Adapter : WkRecordSpreadsheetAdapterBase {
    public override LegacySpreadsheetFormat Format => LegacySpreadsheetFormat.Lotus123;
    public override string ProfileId => "lotus-1-2-3-selected";
    public override string GetProfileId(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x06, 0x04)
            ? "lotus-1-2-3-wk1-records" : LaterRecordSpreadsheetReader.IsLotus(data) ? "lotus-1-2-3-later-records" : "lotus-1-2-3-later-salvage";
    }

    public override int Probe(byte[] data, string? sourceName, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken, out string reason) {
        cancellationToken.ThrowIfCancellationRequested();
        if (LaterRecordSpreadsheetReader.IsLotus(data)) {
            reason = "Lotus WK3/WK4/123 generation-specific BOF payload.";
            return 100;
        }
        if (OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x1A, 0x00)) {
            if (ExtensionIs(sourceName, ".wk3", ".wk4", ".123")) {
                reason = "Later Lotus 1-2-3 BOF record envelope with a corroborating Lotus family extension.";
                return 95;
            }
            reason = "Ambiguous later WK-family BOF record envelope without corroborating Lotus family evidence.";
            return 45;
        }
        if (OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x06, 0x04)) {
            reason = ExtensionIs(sourceName, ".wk1", ".wk2", ".wk3", ".wk4", ".123")
                ? "Exact Lotus WK1 BOF signature with corroborating source extension."
                : "Exact Lotus WK1 BOF signature.";
            return 100;
        }
        if (ExtensionIs(sourceName, ".wk1", ".wk2", ".wk3", ".wk4", ".123")) {
            reason = "Lotus 1-2-3 family source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No Lotus 1-2-3 signature evidence.";
        return 0;
    }

    public override LegacySpreadsheetModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        if (LaterRecordSpreadsheetReader.IsLotus(data)) return LaterRecordSpreadsheetReader.Read(data, limits, cancellationToken, LaterSpreadsheetProfile.Lotus);
        if (OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x06, 0x04)) return ParseWkRecords(data, limits, "Lotus WK1", 0x06, 0x04, cancellationToken);
        return ParseDelimitedSalvage(data, limits,
            "Later Lotus workbook text was salvaged; sheet structure, formulas, formatting, names, comments, and charts require a profile-specific decoder and are reported as loss.", cancellationToken);
    }
}
