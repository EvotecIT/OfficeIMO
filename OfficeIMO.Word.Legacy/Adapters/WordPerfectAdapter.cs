using System;

namespace OfficeIMO.Word.Legacy;

internal sealed class WordPerfectAdapter : LegacyWordAdapterBase {
    public override LegacyWordFormat Format => LegacyWordFormat.WordPerfect;
    public override string ProfileId => "wordperfect-5-6";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        if (OfficeLegacyImportBuffer.StartsWith(data, 0xFF, 0x57, 0x50, 0x43)) {
            reason = "WordPerfect file prefix signature.";
            return 100;
        }
        if (ExtensionIs(sourceName, ".wp", ".wp5", ".wp6", ".wpd")) {
            reason = "WordPerfect-family source extension.";
            return 45;
        }
        reason = "No WordPerfect signature evidence.";
        return 0;
    }

    public override LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        int offset = data.Length >= 8 ? OfficeLegacyImportBuffer.ReadInt32(data, 4) : 0;
        if (offset < 0 || offset >= data.Length) offset = 0;
        LegacyWordModel model = Salvage(data, limits, offset, false, ProfileId,
            "WordPerfect document-area text was recovered conservatively; unsupported prefix packets, formatting codes, notes, tables, graphics, and layout are reported as loss.", cancellationToken);
        model.Metadata["DocumentAreaOffset"] = offset.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (ContainsAscii(data, "MACRO", cancellationToken) || ContainsAscii(data, "PerfectScript", cancellationToken)) {
            model.InertContent |= OfficeLegacyInertContentKind.Macros | OfficeLegacyInertContentKind.EmbeddedCode;
            model.Findings.Add(Inert("WORDPERFECT_ACTIVE_CONTENT_INERT", "Security", "WordPerfect macro or script markers were detected and kept inert."));
        }
        return model;
    }

    private static bool ContainsAscii(byte[] data, string value, System.Threading.CancellationToken cancellationToken) {
        byte[] needle = System.Text.Encoding.ASCII.GetBytes(value);
        for (int index = 0; index <= data.Length - needle.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int matched = 0;
            while (matched < needle.Length && data[index + matched] == needle[matched]) matched++;
            if (matched == needle.Length) return true;
        }
        return false;
    }
}
