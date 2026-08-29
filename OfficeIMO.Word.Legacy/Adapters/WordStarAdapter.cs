using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Legacy;

internal sealed class WordStarAdapter : LegacyWordAdapterBase {
    public override LegacyWordFormat Format => LegacyWordFormat.WordStar;
    public override string ProfileId => "wordstar-3-7";

    public override int Probe(byte[] data, string? sourceName, out string reason) {
        int highAscii = 0;
        int printable = 0;
        for (int index = 0; index < Math.Min(data.Length, 4096); index++) {
            byte value = data[index];
            if ((value & 0x7F) >= 32 && (value & 0x7F) <= 126) printable++;
            if (value >= 0x80 && (value & 0x7F) >= 32 && (value & 0x7F) <= 126) highAscii++;
        }
        if (highAscii >= 8 && printable > 0 && highAscii * 4 >= printable) {
            reason = "WordStar high-bit character flags.";
            return 70;
        }
        if (ExtensionIs(sourceName, ".ws", ".ws3", ".ws4", ".ws5", ".ws6", ".ws7")) {
            reason = "WordStar-family source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No WordStar signature evidence.";
        return 0;
    }

    public override LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        string text = OfficeLegacyImportBuffer.ExtractPrintableText(data, 0, data.Length, limits.MaxTextCharacters, stripHighBit: true, minimumRunLength: 1, cancellationToken: cancellationToken);
        if (string.IsNullOrWhiteSpace(text)) throw new InvalidDataException("WordStar source did not contain recoverable text.");
        var filtered = new StringBuilder(text.Length);
        var model = new LegacyWordModel { Quality = OfficeLegacyImportQuality.Structured };
        foreach (string line in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
            cancellationToken.ThrowIfCancellationRequested();
            if (line.StartsWith(".", StringComparison.Ordinal)) {
                model.Findings.Add(Loss("WORDSTAR_DOT_COMMAND", "Layout", "A WordStar dot command was kept inert and omitted from the editable document."));
                continue;
            }
            filtered.AppendLine(line);
        }
        AddParagraphs(model, filtered.ToString(), limits, cancellationToken);
        model.Findings.Add(Loss("WORDSTAR_FORMATTING_PARTIAL", "Formatting", "WordStar character flags and paragraph boundaries were recovered; unsupported print controls and page layout were omitted."));
        return model;
    }
}
