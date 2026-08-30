namespace OfficeIMO.Excel.Legacy;

internal sealed class QuattroProAdapter : WkRecordSpreadsheetAdapterBase {
    public override LegacySpreadsheetFormat Format => LegacySpreadsheetFormat.QuattroPro;
    public override string ProfileId => "quattro-pro-selected";
    public override string GetProfileId(byte[] data) => OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x20, 0x51)
        ? "quattro-pro-wq1-records" : OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x21, 0x51)
            ? "quattro-pro-wq2-records" : "quattro-pro-wb-qpw-salvage";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        bool wq = OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x20, 0x51) ||
                  OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x21, 0x51);
        bool wb = OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x01, 0x10) ||
                  OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x02, 0x10) ||
                  OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x07, 0x10);
        if (wq || wb) {
            reason = "Quattro Pro WQ/WB BOF payload signature.";
            return 100;
        }
        if (ExtensionIs(sourceName, ".qpw") && OfficeLegacyCompoundInspector.IsValidCompound(data)) {
            reason = "Validated compound workbook with Quattro Pro extension.";
            return 95;
        }
        if (ExtensionIs(sourceName, ".wq1", ".wq2", ".wb1", ".wb2", ".wb3", ".qpw")) {
            reason = "Quattro Pro family source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No Quattro Pro signature evidence.";
        return 0;
    }

    public override LegacySpreadsheetModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        if (OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x20, 0x51)) {
            return ParseWkRecords(data, limits, "Quattro Pro WQ1", 0x20, 0x51, cancellationToken, translateFormulas: false);
        }
        if (OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x21, 0x51)) {
            return ParseWkRecords(data, limits, "Quattro Pro WQ2", 0x21, 0x51, cancellationToken, WkRecordLayout.QuattroWq2, translateFormulas: false);
        }
        LegacySpreadsheetModel model = ParseDelimitedSalvage(data, limits,
            "Quattro Pro compound-workbook text was salvaged; workbook structure, formulas, formatting, comments, and charts were not reconstructed.", cancellationToken);
        model.InertContent |= OfficeLegacyCompoundInspector.Inspect(data, limits, out bool inspectionIncomplete, cancellationToken);
        if (inspectionIncomplete) {
            model.Findings.Add(Loss("LEGACY_COMPOUND_INVENTORY_INCOMPLETE", "Security", "The compound directory could not be inspected within configured safety limits; active-content inventory is indeterminate and no compound stream was activated."));
        }
        if (model.InertContent != OfficeLegacyInertContentKind.None) {
            model.Findings.Add(Inert("QUATTRO_COMPOUND_CONTENT_INERT", "Security", "Active or externally resolved Quattro Pro compound streams were inventoried but never activated, executed, or refreshed."));
        }
        return model;
    }
}
