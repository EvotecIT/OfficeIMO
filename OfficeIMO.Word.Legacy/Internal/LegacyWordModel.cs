using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Word.Legacy;

internal sealed class LegacyWordModel {
    internal List<LegacyWordParagraph> Paragraphs { get; } = new();
    internal Dictionary<string, string> Metadata { get; } = new(System.StringComparer.OrdinalIgnoreCase);
    internal List<OfficeCompatibilityFinding> Findings { get; } = new();
    internal OfficeLegacyImportQuality Quality { get; set; } = OfficeLegacyImportQuality.Salvage;
    internal OfficeLegacyInertContentKind InertContent { get; set; }
}

internal sealed class LegacyWordParagraph {
    internal LegacyWordParagraph(string text, bool isList = false, int listLevel = 0) {
        Text = text;
        IsList = isList;
        ListLevel = listLevel;
    }

    internal string Text { get; }
    internal bool IsList { get; }
    internal int ListLevel { get; }
}

internal interface ILegacyWordAdapter {
    LegacyWordFormat Format { get; }
    string ProfileId { get; }
    int Probe(byte[] data, string? sourceName, out string reason);
    LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken);
}
