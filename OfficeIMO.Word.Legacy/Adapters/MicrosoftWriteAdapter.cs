namespace OfficeIMO.Word.Legacy;

internal sealed class MicrosoftWriteAdapter : LegacyWordAdapterBase {
    public override LegacyWordFormat Format => LegacyWordFormat.MicrosoftWrite;
    public override string ProfileId => "microsoft-write-wri-salvage";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        bool sharedHeader = data.Length > 96 && (data[0] == 0x31 || data[0] == 0x32) && data[1] == 0xBE && data[5] == 0xAB;
        if (sharedHeader && data[96] != 0) {
            reason = "Microsoft Write binary header and nonzero Write discriminator.";
            return 100;
        }
        if (ExtensionIs(sourceName, ".wri")) {
            reason = "Microsoft Write source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No Microsoft Write signature evidence.";
        return 0;
    }

    public override LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) =>
        Salvage(data, limits, System.Math.Min(128, data.Length), false, ProfileId,
            "Microsoft Write text and paragraph boundaries were salvaged; formatting runs, objects, headers, footers, and page layout were omitted.", cancellationToken);
}
