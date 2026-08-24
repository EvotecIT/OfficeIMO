using System.Text;
using System.Text.Json;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusReportWriter {
    public static void Write(CorpusReport report, string jsonPath, string markdownPath) {
        Directory.CreateDirectory(Path.GetDirectoryName(jsonPath) ?? ".");
        Directory.CreateDirectory(Path.GetDirectoryName(markdownPath) ?? ".");
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(report, CorpusJson.Options) + Environment.NewLine, new UTF8Encoding(false));
        File.WriteAllText(markdownPath, RenderMarkdown(report), new UTF8Encoding(false));
    }

    internal static string RenderMarkdown(CorpusReport report) {
        var text = new StringBuilder();
        text.AppendLine("# Real-world corpus evidence").AppendLine();
        text.AppendLine("> This is bounded discovery evidence, not a claim that every document, producer, or feature works. " +
                        "The source is a convenience corpus, files are not independent random observations, and this lane does not assess visual fidelity.").AppendLine();
        text.AppendLine("## Provenance").AppendLine();
        text.AppendLine("| Field | Value |");
        text.AppendLine("| --- | --- |");
        text.Append("| Measurement | ").Append(E(report.MeasurementStatus)).AppendLine(" |");
        text.Append("| Corpus | ").Append(E(report.Provenance.CorpusId)).AppendLine(" |");
        text.Append("| Source | ").Append(E(report.Provenance.SourceUri ?? "not recorded")).AppendLine(" |");
        text.Append("| Archive SHA-256 | `").Append(E(report.Provenance.ArchiveSha256 ?? "not recorded")).AppendLine("` |");
        text.Append("| Started UTC | ").Append(report.StartedUtc.ToString("O")).AppendLine(" |");
        text.Append("| Completed UTC | ").Append(report.CompletedUtc.ToString("O")).AppendLine(" |");
        text.Append("| Runtime | ").Append(E(report.Environment.Framework)).Append(" on ").Append(E(report.Environment.OperatingSystem)).AppendLine(" |");
        text.AppendLine().AppendLine("## Sample and outcomes").AppendLine();
        text.AppendLine("Selection is deterministic: unique content is ordered by SHA-256 and sampled round-robin across the requested format strata. " +
                        "A stratum with fewer eligible files is reported as underfilled; no other format is relabeled to fill its denominator.").AppendLine();
        text.AppendLine("| Format | Eligible unique | Requested maximum | Selected | Completed | With warnings | With errors | Policy rejected | Failed | Timed out | Underfilled |");
        text.AppendLine("| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |");
        foreach (CorpusStratum stratum in report.Strata) {
            text.Append("| ").Append(stratum.Format).Append(" | ").Append(stratum.EligibleUnique)
                .Append(" | ").Append(stratum.RequestedMaximum).Append(" | ").Append(stratum.Selected)
                .Append(" | ").Append(stratum.Completed).Append(" | ").Append(stratum.CompletedWithWarnings)
                .Append(" | ").Append(stratum.CompletedWithErrors).Append(" | ").Append(stratum.RejectedByPolicy)
                .Append(" | ").Append(stratum.Failed).Append(" | ").Append(stratum.TimedOut)
                .Append(" | ").Append(stratum.CorpusUnderfilled ? "yes" : "no").AppendLine(" |");
        }
        text.AppendLine().AppendLine("## Inventory accounting").AppendLine();
        text.AppendLine("| Discovered | Oversize | Classification failed | Classification timed out | Duplicate content | Eligible unique | Selected |");
        text.AppendLine("| ---: | ---: | ---: | ---: | ---: | ---: | ---: |");
        text.Append("| ").Append(report.Totals.Discovered).Append(" | ").Append(report.Totals.Oversize)
            .Append(" | ").Append(report.Totals.ClassificationFailed).Append(" | ").Append(report.Totals.ClassificationTimedOut)
            .Append(" | ").Append(report.Totals.DuplicateContent).Append(" | ").Append(report.Totals.EligibleUnique)
            .Append(" | ").Append(report.Totals.Selected).AppendLine(" |");

        CorpusFileRecord[] observations = report.Files.Where(file => file.Selected && file.Outcome != CorpusOutcomes.Completed).ToArray();
        text.AppendLine().AppendLine("## Observations requiring interpretation").AppendLine();
        if (observations.Length == 0) {
            text.AppendLine("No selected file produced warnings, errors, process failures, or timeouts in this run.").AppendLine();
        } else {
            text.AppendLine("| SHA-256 | Source | Format | Outcome | Diagnostics | Exception |");
            text.AppendLine("| --- | --- | --- | --- | --- | --- |");
            foreach (CorpusFileRecord file in observations) {
                text.Append("| `").Append(ShortHash(file.Sha256)).Append("` | ").Append(E(file.SourceName ?? "withheld"))
                    .Append(" | ").Append(file.ContentKind).Append(" | ").Append(E(file.Outcome))
                    .Append(" | ").Append(E(string.Join(", ", file.DiagnosticCodes)))
                    .Append(" | ").Append(E(file.ExceptionType ?? string.Empty)).AppendLine(" |");
            }
            text.AppendLine();
        }
        text.AppendLine("## Interpretation boundary").AppendLine();
        text.AppendLine("A target of 100 unique files per format is large enough to expose many recurring defects while remaining practical for a monthly isolated-process run. " +
                        "The familiar rule-of-three would put an approximate 95% upper bound near 3% after 100 independent observations with zero failures, " +
                        "but these corpus files are not an independent random sample. That calculation explains the sample budget; it is not a reliability guarantee.").AppendLine();
        text.AppendLine("The lane proves only that OfficeIMO content-detected and attempted its normalized read contract for the recorded hashes under the recorded limits and runtime. " +
                        "It does not prove rendering fidelity, editing round trips, semantic preservation of every feature, safety of opening files in another application, or support for unmeasured files. " +
                        "Actionable findings should be minimized into provenance-tracked fixtures and moved to the owning format test suite.");
        return text.ToString().Replace("\r\n", "\n");
    }

    private static string ShortHash(string? value) => string.IsNullOrEmpty(value) ? "unavailable" : value[..Math.Min(16, value.Length)];
    private static string E(string value) {
        var escaped = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (char.IsControl(character)) {
                escaped.Append(' ');
            } else if (char.IsPunctuation(character) || char.IsSymbol(character)) {
                escaped.Append("&#").Append((int)character).Append(';');
            } else {
                escaped.Append(character);
            }
        }
        return escaped.ToString();
    }
}
