using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Legacy;

internal sealed class AmiProAdapter : LegacyWordAdapterBase {
    public override LegacyWordFormat Format => LegacyWordFormat.AmiPro;
    public override string ProfileId => "ami-pro-sam";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        string prefix = Encoding.ASCII.GetString(data, 0, Math.Min(data.Length, 128));
        if (prefix.IndexOf("[ver]", StringComparison.OrdinalIgnoreCase) >= 0 || prefix.IndexOf("[ami", StringComparison.OrdinalIgnoreCase) >= 0) {
            reason = "Ami Pro tagged-text header.";
            return 95;
        }
        if (ExtensionIs(sourceName, ".sam")) {
            reason = "Ami Pro SAM source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No Ami Pro signature evidence.";
        return 0;
    }

    public override LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        string text = OfficeLegacyImportBuffer.ExtractPrintableText(data, 0, data.Length, limits.MaxTextCharacters, false, 1, cancellationToken);
        if (string.IsNullOrWhiteSpace(text)) throw new InvalidDataException("Ami Pro source did not contain recoverable text.");
        var visible = new StringBuilder(text.Length);
        var model = new LegacyWordModel { Quality = OfficeLegacyImportQuality.Structured };
        foreach (string raw in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
            cancellationToken.ThrowIfCancellationRequested();
            string line = raw.TrimEnd();
            if (line.StartsWith("[", StringComparison.Ordinal) || line.StartsWith("@", StringComparison.Ordinal)) continue;
            visible.AppendLine(line);
        }
        AddParagraphs(model, visible.ToString(), limits, cancellationToken);
        model.Findings.Add(Loss("AMIPRO_TAGS_PARTIAL", "Formatting", "Ami Pro tagged text and paragraph boundaries were recovered; unsupported styles, frames, notes, tables, and graphics were omitted."));
        return model;
    }
}
