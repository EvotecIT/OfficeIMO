using System.Globalization;
using System.Text;
using OfficeIMO.Provenance;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Workflows;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>Finds invisible and context-sensitive characters in text and writes a copy without the ones the visitor selects.</summary>
internal static class TextTool {
    private sealed record Review(OfficeTextIntegrityReview Result, string Text, string SourceName);

    private static OfficeTextIntegrityOptions Limits() => new() { MaxEncodedBytes = 1024 * 1024, MaxCharacters = 262_144, MaxFindings = 512 };

    internal static ToolResultDocument Run(ToolSession session, string action, ToolOptions options) => action switch {
        "inspect" => Inspect(session, options),
        "remove" => Remove(session, options),
        _ => throw new NotSupportedException($"The hidden characters tool has no '{action}' action.")
    };

    private static ToolResultDocument Inspect(ToolSession session, ToolOptions options) {
        OfficeTextIntegrityReview result;
        string sourceName;
        if (options.Text("source") == "file") {
            SelectedDocument file = session.Input();
            sourceName = file.Name;
            result = OfficeTextIntegrityReview.Inspect(file.Bytes, Limits(), file.Name);
        } else {
            string text = options.Text("text");
            if (text.Length == 0) throw new InvalidOperationException("Paste some text or open a text file first.");
            sourceName = "text.txt";
            result = OfficeTextIntegrityReview.Inspect(text, Limits(), sourceName);
        }
        var review = new Review(result, result.Text, sourceName);
        session.LastReview = review;
        ToolArtifact[] artifacts = session.ReplaceArtifacts((ReportArtifact(review, null), "report"));
        IReadOnlyList<OfficeTextIntegrityFinding> findings = result.Report.Findings;
        int risky = findings.Count(static finding => finding.Risk == OfficeTextIntegrityRisk.PotentiallyDangerous);
        ToolVerdict verdict = findings.Count == 0
            ? new ToolVerdict(ToolTone.Good, "No hidden characters found", "The text has no invisible or direction-changing characters.")
            : new ToolVerdict(risky > 0 ? ToolTone.Warn : ToolTone.Info,
                "Found " + ToolFormat.Count(findings.Count, "hidden or special character"),
                risky > 0
                    ? ToolFormat.Count(risky, "of them can", "of them can") + " change how the text displays or how programs read it. They are highlighted below."
                    : "They are usually intentional, such as emoji joiners or non-breaking spaces. Remove only what you don't expect.");
        return new ToolResultDocument(true, verdict,
            [new ToolFact("File", sourceName), new ToolFact("Encoding", result.EncodingName.ToUpperInvariant()),
             new ToolFact("Characters", result.Text.Length.ToString("N0", CultureInfo.InvariantCulture)),
             new ToolFact("Findings", findings.Count.ToString(CultureInfo.InvariantCulture), findings.Count > 0 ? ToolTone.Warn : null)],
            Items(findings, null), artifacts, new ToolPreview("text", Text: result.Text), 0);
    }

    private static ToolResultDocument Remove(ToolSession session, ToolOptions options) {
        Review review = session.LastReview as Review ?? throw new InvalidOperationException("Inspect the text before removing characters.");
        if (options.Text("source") != "file" && options.Text("text") != review.Text) {
            throw new InvalidOperationException("The text changed. Inspect it again before removing characters.");
        }
        int count = review.Result.Report.Findings.Count;
        int[] selected = options.List("remove")
            .Select(static id => int.TryParse(id.TrimStart('f'), NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) ? index : -1)
            .Where(index => index >= 0 && index < count)
            .Distinct()
            .ToArray();
        if (selected.Length == 0) throw new InvalidOperationException("Select at least one character to remove.");
        byte[] output = review.Result.ExportSelected(review.Text, selected);
        string outputName = Path.GetFileNameWithoutExtension(review.SourceName) + "-cleaned" + (Path.GetExtension(review.SourceName) is { Length: > 0 } extension ? extension : ".txt");
        ToolArtifact[] artifacts = session.ReplaceArtifacts(
            (new BrowserConversionArtifact(output, outputName, "text/plain;charset=" + review.Result.EncodingName), "primary"),
            (ReportArtifact(review, selected), "report"));
        return new ToolResultDocument(true,
            new ToolVerdict(ToolTone.Good, "Clean copy ready · " + ToolFormat.Count(selected.Length, "character") + " removed",
                count > selected.Length ? ToolFormat.Count(count - selected.Length, "character") + " you left unselected stayed in place." : "Everything that was found is gone."),
            [new ToolFact("Removed", selected.Length.ToString(CultureInfo.InvariantCulture), ToolTone.Good), new ToolFact("Size", ToolFormat.Bytes(output.LongLength))],
            // The preview keeps the original text so each marker sits where it was; removed ones are shown struck through.
            Items(review.Result.Report.Findings, selected.ToHashSet()), artifacts, new ToolPreview("text", 0, Text: review.Text), 0);
    }

    private static ToolItem[] Items(IReadOnlyList<OfficeTextIntegrityFinding> findings, IReadOnlySet<int>? removed) =>
        findings.Select((finding, index) => new ToolItem(
            "f" + index.ToString(CultureInfo.InvariantCulture),
            KindName(finding.Kind) + " (" + finding.UnicodeNotation + ")",
            RiskText(finding.Risk),
            removed is null ? ToolState.Found : removed.Contains(index) ? ToolState.Removed : ToolState.Kept,
            RiskGroup(finding.Risk),
            Selectable: removed is null,
            Selected: false,
            Start: finding.TextOffset,
            Length: finding.TextLength)).ToArray();

    private static BrowserConversionArtifact ReportArtifact(Review review, IEnumerable<int>? selected) =>
        new(Encoding.UTF8.GetBytes(selected is null
                ? OfficeTextIntegrityReportSerializer.Serialize(review.Result, review.Text, review.SourceName)
                : OfficeTextIntegrityReportSerializer.Serialize(review.Result, review.Text, review.SourceName, selected)),
            "hidden-characters-report.json",
            "application/json");

    private static string KindName(OfficeTextIntegrityFindingKind kind) => kind switch {
        OfficeTextIntegrityFindingKind.EmbeddedByteOrderMark => "Byte-order mark",
        OfficeTextIntegrityFindingKind.ZeroWidthSpace => "Zero-width space",
        OfficeTextIntegrityFindingKind.ZeroWidthNonJoiner => "Zero-width non-joiner",
        OfficeTextIntegrityFindingKind.ZeroWidthJoiner => "Zero-width joiner",
        OfficeTextIntegrityFindingKind.WordJoiner => "Word joiner",
        OfficeTextIntegrityFindingKind.BidirectionalControl => "Text direction control",
        OfficeTextIntegrityFindingKind.UnicodeTag => "Hidden tag character",
        OfficeTextIntegrityFindingKind.VariationSelector => "Variation selector",
        OfficeTextIntegrityFindingKind.TypographicSpace => "Special space",
        OfficeTextIntegrityFindingKind.InvisibleFormatCharacter => "Invisible formatting character",
        OfficeTextIntegrityFindingKind.ControlCharacter => "Control character",
        OfficeTextIntegrityFindingKind.UnpairedSurrogate => "Broken character",
        _ => ToolFormat.Words(kind.ToString())
    };

    private static string RiskText(OfficeTextIntegrityRisk risk) => risk switch {
        OfficeTextIntegrityRisk.PotentiallyDangerous => "Can change how the text displays or how programs read it. Usually safe to remove.",
        OfficeTextIntegrityRisk.ContextDependent => "Often intentional, for example inside emoji or in some languages. Remove it only if you didn't expect it.",
        _ => "Usually intentional typography."
    };

    private static string RiskGroup(OfficeTextIntegrityRisk risk) => risk switch {
        OfficeTextIntegrityRisk.PotentiallyDangerous => "Can change meaning",
        OfficeTextIntegrityRisk.ContextDependent => "Often intentional",
        _ => "Typography"
    };
}
