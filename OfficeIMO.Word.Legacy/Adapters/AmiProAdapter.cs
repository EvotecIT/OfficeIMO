using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Word.Legacy;

internal sealed class AmiProAdapter : LegacyWordAdapterBase {
    public override LegacyWordFormat Format => LegacyWordFormat.AmiPro;
    public override string ProfileId => "ami-pro-sam-v4";
    public override string GetProfileId(byte[] data) => TryGetVersion(data, out int version) && version == 4 ? ProfileId : "ami-pro-sam-unsupported-salvage";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        if (TryGetVersion(data, out int version) && version == 4) {
            reason = "Ami Pro SAM version 4 header.";
            return 95;
        }
        if (TryGetVersion(data, out version)) {
            reason = $"Ami Pro SAM version {version} header is outside the structured SAM4 profile; use FormatHint for bounded salvage.";
            return 40;
        }
        if (ExtensionIs(sourceName, ".sam")) {
            reason = "Ami Pro SAM source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No Ami Pro signature evidence.";
        return 0;
    }

    public override LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        if (TryGetVersion(data, out int version) && version == 4) return new AmiProSamParser(data, limits, cancellationToken).Parse();
        return Salvage(data, limits, 0, stripHighBit: false, "ami-pro-sam-unsupported-salvage",
            "Ami Pro tagged text outside the SAM version 4 structured profile was salvaged; styles, frames, notes, images, and layout are not claimed.", cancellationToken);
    }

    private static bool TryGetVersion(byte[] data, out int version) {
        version = 0;
        string prefix = Encoding.ASCII.GetString(data, 0, Math.Min(data.Length, 4096)).Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = prefix.Split('\n');
        for (int index = 0; index + 1 < lines.Length; index++) {
            if (!string.Equals(lines[index].Trim('\uFEFF', '\0', ' ', '\t'), "[ver]", StringComparison.OrdinalIgnoreCase)) continue;
            return int.TryParse(lines[index + 1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out version);
        }
        return false;
    }
}
