namespace OfficeIMO.Excel.Legacy;

internal sealed class MicrosoftWorksSpreadsheetAdapter : WkRecordSpreadsheetAdapterBase {
    public override LegacySpreadsheetFormat Format => LegacySpreadsheetFormat.MicrosoftWorks;
    public override string ProfileId => "microsoft-works-spreadsheet-selected";
    public override string GetProfileId(byte[] data, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x04, 0x04)
            ? "microsoft-works-wks-dos-records" : OfficeLegacyCompoundInspector.IsValidCompound(data, cancellationToken)
                ? "microsoft-works-xlr-compound-salvage" : "microsoft-works-spreadsheet-binary-salvage";
    }

    public override int Probe(byte[] data, string? sourceName, System.Threading.CancellationToken cancellationToken, out string reason) {
        cancellationToken.ThrowIfCancellationRequested();
        if (OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x04, 0x04)) {
            reason = ExtensionIs(sourceName, ".wks")
                ? "Exact Microsoft Works WKS BOF signature with corroborating source extension."
                : "Exact Microsoft Works WKS BOF signature.";
            return 100;
        }
        if (OfficeLegacyImportBuffer.StartsWith(data, 0xFF, 0x00, 0x02)) {
            if (ExtensionIs(sourceName, ".wks")) {
                reason = "Microsoft Works 3.x spreadsheet prefix with corroborating Works extension.";
                return 95;
            }
            reason = "Ambiguous three-byte Works-family prefix without corroborating source evidence.";
            return 45;
        }
        if (ExtensionIs(sourceName, ".xlr") && OfficeLegacyCompoundInspector.IsValidCompound(data, cancellationToken)) {
            reason = "OLE compound workbook with Microsoft Works spreadsheet extension.";
            return 95;
        }
        if (ExtensionIs(sourceName, ".wks", ".xlr")) {
            reason = "Microsoft Works spreadsheet source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No Microsoft Works spreadsheet signature evidence.";
        return 0;
    }

    public override LegacySpreadsheetModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        if (OfficeLegacyImportBuffer.StartsWith(data, 0x00, 0x00, 0x02, 0x00, 0x04, 0x04)) return ParseWkRecords(data, limits, "Microsoft Works WKS", 0x04, 0x04, cancellationToken);
        LegacySpreadsheetModel model = ParseDelimitedSalvage(data, limits,
            "Microsoft Works spreadsheet text was salvaged; sheet structure, formulas, formatting, comments, and charts were not reconstructed.", cancellationToken);
        model.InertContent |= OfficeLegacyCompoundInspector.Inspect(data, limits, out bool inspectionIncomplete, cancellationToken);
        if (inspectionIncomplete) {
            model.Findings.Add(Loss("LEGACY_COMPOUND_INVENTORY_INCOMPLETE", "Security", "The compound directory could not be inspected within configured safety limits; active-content inventory is indeterminate and no compound stream was activated."));
        }
        if (model.InertContent != OfficeLegacyInertContentKind.None) {
            model.Findings.Add(Inert("WORKS_SHEET_ACTIVE_CONTENT_INERT", "Security", "Active or externally resolved Works compound streams were inventoried but never activated, executed, or refreshed."));
        }
        return model;
    }
}
