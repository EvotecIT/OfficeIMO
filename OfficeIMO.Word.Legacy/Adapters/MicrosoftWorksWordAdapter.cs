namespace OfficeIMO.Word.Legacy;

internal sealed class MicrosoftWorksWordAdapter : LegacyWordAdapterBase {
    public override LegacyWordFormat Format => LegacyWordFormat.MicrosoftWorks;
    public override string ProfileId => "microsoft-works-word-2-8-salvage";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        if (ExtensionIs(sourceName, ".wps") && OfficeLegacyCompoundInspector.IsValidCompound(data)) {
            reason = "OLE compound document with Microsoft Works word-processing extension.";
            return 90;
        }
        if (data.Length >= 2 && data[0] < 6 && data[1] == 0xFE) {
            reason = "Microsoft Works 2.x word-processing header.";
            return 95;
        }
        if (ExtensionIs(sourceName, ".wps")) {
            reason = "Microsoft Works word-processing extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No Microsoft Works word-processing signature evidence.";
        return 0;
    }

    public override LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        LegacyWordModel model = Salvage(data, limits, 0, false, ProfileId,
            "Microsoft Works text was salvaged; unsupported paragraph properties, fields, notes, tables, images, and layout were omitted.", cancellationToken);
        model.InertContent |= OfficeLegacyCompoundInspector.Inspect(data, limits, out bool inspectionIncomplete, cancellationToken);
        if (inspectionIncomplete) {
            model.Findings.Add(Loss("LEGACY_COMPOUND_INVENTORY_INCOMPLETE", "Security", "The compound directory could not be inspected within configured safety limits; active-content inventory is indeterminate and no compound stream was activated."));
        }
        if (model.InertContent != OfficeLegacyInertContentKind.None) {
            model.Findings.Add(Inert("WORKS_ACTIVE_CONTENT_INERT", "Security", "Active or externally resolved Works compound streams were inventoried but never activated, executed, or resolved."));
        }
        return model;
    }
}
