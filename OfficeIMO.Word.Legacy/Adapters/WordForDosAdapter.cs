namespace OfficeIMO.Word.Legacy;

internal sealed class WordForDosAdapter : LegacyWordAdapterBase {
    public override LegacyWordFormat Format => LegacyWordFormat.WordForDos;
    public override string ProfileId => "microsoft-word-dos-4-6";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        bool sharedHeader = data.Length > 96 && (data[0] == 0x31 || data[0] == 0x32) && data[1] == 0xBE && data[5] == 0xAB;
        if (sharedHeader && data[96] == 0) {
            reason = "Selected Word for DOS binary header and zero DOS discriminator.";
            return 100;
        }
        if (ExtensionIs(sourceName, ".doc") && data.Length > 0 && !OfficeLegacyImportBuffer.StartsWith(data, 0xD0, 0xCF, 0x11, 0xE0)) {
            reason = "Weak extension-assisted Word for DOS candidate.";
            return 20;
        }
        reason = "No selected Word for DOS signature evidence.";
        return 0;
    }

    public override LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) =>
        Salvage(data, limits, System.Math.Min(128, data.Length), false, ProfileId,
            "Selected Word for DOS text and paragraph boundaries were salvaged; formatting, annotations, objects, and page layout were omitted.", cancellationToken);
}
