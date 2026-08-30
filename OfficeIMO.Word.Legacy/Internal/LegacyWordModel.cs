using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Legacy;

internal sealed class LegacyWordModel {
    internal List<LegacyWordParagraph> Paragraphs { get; } = new();
    internal List<LegacyWordNote> Notes { get; } = new();
    internal List<LegacyWordResource> Resources { get; } = new();
    internal Dictionary<string, string> Metadata { get; } = new(System.StringComparer.OrdinalIgnoreCase);
    internal List<OfficeCompatibilityFinding> Findings { get; } = new();
    internal OfficeLegacyImportQuality Quality { get; set; } = OfficeLegacyImportQuality.Salvage;
    internal OfficeLegacyInertContentKind InertContent { get; set; }
}

internal sealed class LegacyWordParagraph {
    internal LegacyWordParagraph() { }

    internal LegacyWordParagraph(string text, bool isList = false, int listLevel = 0) {
        Runs.Add(new LegacyWordRun(text));
        IsList = isList;
        ListLevel = listLevel;
    }

    internal LegacyWordParagraph(IEnumerable<LegacyWordRun> runs) => Runs.AddRange(runs);

    internal List<LegacyWordRun> Runs { get; } = new();
    internal string Text => string.Concat(Runs.Select(static run => run.Text));
    internal bool IsList { get; set; }
    internal int ListLevel { get; set; }
    internal WordParagraphAlignment? Alignment { get; set; }
    internal bool PageBreakBefore { get; set; }
    internal bool KeepWithNext { get; set; }
    internal bool KeepLinesTogether { get; set; }
    internal double? LineSpacingPoints { get; set; }
    internal double? SpacingBeforePoints { get; set; }
    internal double? SpacingAfterPoints { get; set; }
    internal string? StyleName { get; set; }
}

internal sealed class LegacyWordRun {
    internal LegacyWordRun(string text) => Text = text;
    internal string Text { get; }
    internal bool Bold { get; set; }
    internal bool Italic { get; set; }
    internal bool Strike { get; set; }
    internal WordUnderlineStyle? Underline { get; set; }
    internal WordVerticalTextPosition? VerticalPosition { get; set; }
    internal int? FontSizePoints { get; set; }
    internal string? FontFamily { get; set; }
    internal string? ColorHex { get; set; }
}

internal sealed class LegacyWordNote {
    internal LegacyWordNote(LegacyWordNoteKind kind, string text) { Kind = kind; Text = text; }
    internal LegacyWordNoteKind Kind { get; }
    internal string Text { get; }
}

internal sealed class LegacyWordResource {
    internal LegacyWordResource(string kind, string reference) { Kind = kind; Reference = reference; }
    internal string Kind { get; }
    internal string Reference { get; }
}

internal interface ILegacyWordAdapter {
    LegacyWordFormat Format { get; }
    string ProfileId { get; }
    string GetProfileId(byte[] data, CancellationToken cancellationToken);
    int Probe(byte[] data, string? sourceName, CancellationToken cancellationToken, out string reason);
    LegacyWordModel Parse(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken);
}
