namespace OfficeIMO.Word.Legacy;

internal sealed class LotusWordProAdapter : LegacyWordAdapterBase {
    public override LegacyWordFormat Format => LegacyWordFormat.LotusWordPro;
    public override string ProfileId => "lotus-word-pro-lwp-salvage";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        if (ExtensionIs(sourceName, ".lwp") && OfficeLegacyCompoundInspector.IsValidCompound(data)) {
            reason = "OLE compound document with Lotus Word Pro extension.";
            return 90;
        }
        if (ExtensionIs(sourceName, ".lwp")) {
            reason = "Lotus Word Pro source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No Lotus Word Pro signature evidence.";
        return 0;
    }

    public override LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        LegacyWordModel model = Salvage(data, limits, 0, false, ProfileId,
            "Lotus Word Pro compound-stream text was salvaged; document structure, styles, notes, tables, graphics, and layout were not reconstructed.", cancellationToken);
        model.InertContent |= OfficeLegacyCompoundInspector.Inspect(data, limits, out bool inspectionIncomplete, cancellationToken);
        if (inspectionIncomplete) {
            model.Findings.Add(Loss("LEGACY_COMPOUND_INVENTORY_INCOMPLETE", "Security", "The compound directory could not be inspected within configured safety limits; active-content inventory is indeterminate and no compound stream was activated."));
        }
        if (model.InertContent != OfficeLegacyInertContentKind.None) {
            model.Findings.Add(Inert("WORDPRO_ACTIVE_CONTENT_INERT", "Security", "Active or externally resolved Word Pro compound streams were inventoried but never activated, executed, or resolved."));
        }
        return model;
    }
}
