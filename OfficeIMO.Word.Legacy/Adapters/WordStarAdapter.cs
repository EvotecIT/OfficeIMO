using System;
using System.Threading;

namespace OfficeIMO.Word.Legacy;

internal sealed class WordStarAdapter : LegacyWordAdapterBase {
    public override LegacyWordFormat Format => LegacyWordFormat.WordStar;
    public override string ProfileId => "wordstar-3-7-character-stream";
    public override string GetProfileId(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken) =>
        HasCoherentGrammar(data, out _, cancellationToken) ? ProfileId : "wordstar-family-salvage";

    public override int Probe(byte[] data, string? sourceName, OfficeLegacyImportLimits limits, CancellationToken cancellationToken, out string reason) {
        if (HasCoherentGrammar(data, out bool malformed, cancellationToken)) {
            reason = "Coherent WordStar character-stream controls, paragraph terminator, and EOF grammar.";
            return 85;
        }
        if (malformed) {
            reason = "Malformed WordStar symmetrical-sequence grammar.";
            return 0;
        }
        if (ExtensionIs(sourceName, ".ws", ".ws3", ".ws4", ".ws5", ".ws6", ".ws7")) {
            reason = "WordStar-family source extension only; use FormatHint when the family is independently known.";
            return 35;
        }
        reason = "No WordStar signature evidence.";
        return 0;
    }

    public override LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, System.Threading.CancellationToken cancellationToken) {
        if (HasCoherentGrammar(data, out bool malformed, cancellationToken)) return new WordStarStructuredParser(data, limits, cancellationToken).Parse();
        if (malformed) throw new System.IO.InvalidDataException("Malformed WordStar symmetrical-sequence grammar.");
        return Salvage(data, limits, 0, stripHighBit: true, "wordstar-family-salvage",
            "WordStar family text was salvaged because a coherent WordStar 3-7 control/paragraph/EOF grammar was not present; formatting, notes, images, and layout are not claimed.", cancellationToken);
    }

    private static bool HasCoherentGrammar(byte[] data, out bool malformed, System.Threading.CancellationToken cancellationToken) {
        malformed = false;
        bool eof = false;
        bool hardParagraph = false;
        int formattingControls = 0;
        int highBitPrintable = 0;
        int printable = 0;
        bool validSequence = false;
        int end = data.Length;
        for (int index = 0; index < end; index++) {
            if ((index & 0xFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            byte value = data[index];
            if (value == 0x1A) { eof = true; break; }
            if (value == 0x8D && index + 1 < end && data[index + 1] == 0x0A) {
                index++;
                continue;
            }
            if ((value & 0x7F) >= 0x20 && (value & 0x7F) <= 0x7E) {
                printable++;
                if (value >= 0x80) highBitPrintable++;
            }
            if (value == 0x0D || value == 0x0A) hardParagraph = true;
            if (value == 0x02 || value == 0x13 || value == 0x14 || value == 0x16 || value == 0x18 || value == 0x19) formattingControls++;
            if (value != 0x1D) continue;
            if (index + 4 > data.Length) { malformed = true; return false; }
            int count = data[index + 1] | (data[index + 2] << 8);
            int totalLength = count + 3;
            if (totalLength < 7 || index > data.Length - totalLength) { malformed = true; return false; }
            int suffix = index + totalLength - 3;
            if (data[suffix] != data[index + 1] || data[suffix + 1] != data[index + 2] || data[suffix + 2] != 0x1D) { malformed = true; return false; }
            validSequence = true;
            index += totalLength - 1;
        }
        bool highBitGrammar = highBitPrintable >= 4 && printable > 0 && highBitPrintable * 5 >= printable;
        return eof && hardParagraph && printable > 0 && (validSequence || formattingControls >= 2 || highBitGrammar);
    }
}
