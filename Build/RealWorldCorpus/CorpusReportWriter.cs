using System.Text;
using System.Text.Json;
using System.Globalization;

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
        text.AppendLine().AppendLine("## Recorded limits").AppendLine();
        text.AppendLine("| Input bytes | Traversal entries | Per-file seconds | Workers | Samples per format | Total samples |");
        text.AppendLine("| ---: | ---: | ---: | ---: | ---: | ---: |");
        text.Append("| ").Append(report.Configuration.MaxFileBytes)
            .Append(" | ").Append(report.Configuration.MaxTraversalEntries)
            .Append(" | ").Append(report.Configuration.TimeoutSeconds)
            .Append(" | ").Append(report.Configuration.Parallelism)
            .Append(" | ").Append(report.Configuration.MaxPerFormat)
            .Append(" | ").Append(report.Configuration.MaxTotal)
            .AppendLine(" |");
        text.AppendLine();
        text.AppendLine("| Detection mode | Inspect containers | Detection probe bytes | Detection container entries | Read characters | Read table rows | Compute hashes |");
        text.AppendLine("| --- | --- | ---: | ---: | ---: | ---: | --- |");
        text.Append("| ").Append(report.Configuration.ReaderPolicy.DetectionMode)
            .Append(" | ").Append(report.Configuration.ReaderPolicy.InspectContainers ? "yes" : "no")
            .Append(" | ").Append(report.Configuration.ReaderPolicy.DetectionMaxProbeBytes)
            .Append(" | ").Append(report.Configuration.ReaderPolicy.DetectionMaxContainerEntries)
            .Append(" | ").Append(report.Configuration.ReaderPolicy.ReadMaxCharacters)
            .Append(" | ").Append(report.Configuration.ReaderPolicy.ReadMaxTableRows)
            .Append(" | ").Append(report.Configuration.ReaderPolicy.ComputeHashes ? "yes" : "no")
            .AppendLine(" |");
        text.AppendLine();
        text.AppendLine("| Package bytes | Package parts | Expanded part bytes | Package XML characters | Total expanded bytes | Compression ratio |");
        text.AppendLine("| ---: | ---: | ---: | ---: | ---: | ---: |");
        text.Append("| ").Append(report.Configuration.PackagePolicy.MaxPackageBytes)
            .Append(" | ").Append(report.Configuration.PackagePolicy.MaxPartCount)
            .Append(" | ").Append(report.Configuration.PackagePolicy.MaxPartUncompressedBytes)
            .Append(" | ").Append(report.Configuration.PackagePolicy.MaxXmlCharactersInPart)
            .Append(" | ").Append(report.Configuration.PackagePolicy.MaxTotalUncompressedBytes)
            .Append(" | ").Append(report.Configuration.PackagePolicy.MaxCompressionRatio.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .AppendLine(" |");
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

        text.AppendLine().AppendLine("## PDF quality depth").AppendLine();
        text.AppendLine("Selected PDF files additionally exercise the canonical PDF inspection, semantic recovery, font, image, managed first-page render, mutation-planning, and declared-compliance claim gates in the same isolated worker.").AppendLine();
        text.AppendLine("| Selected PDFs | Deep stages passed | Deep stages | Rendered pages | Mutation plans | Declared claims | Claimable claims | ");
        text.AppendLine("| ---: | ---: | ---: | ---: | ---: | ---: | ---: |");
        text.Append("| ").Append(report.Totals.PdfSelected)
            .Append(" | ").Append(report.Totals.PdfDeepStagesPassed)
            .Append(" | ").Append(report.Totals.PdfDeepStages)
            .Append(" | ").Append(report.Totals.PdfRenderedPages)
            .Append(" | ").Append(report.Totals.PdfMutationPlans)
            .Append(" | ").Append(report.Totals.PdfDeclaredComplianceClaims)
            .Append(" | ").Append(report.Totals.PdfClaimableComplianceClaims).AppendLine(" |");

        CorpusFileRecord[] observations = report.Files.Where(file => file.Selected &&
            (file.Outcome != CorpusOutcomes.Completed || file.PdfEvidence?.AllStagesSucceeded == false)).ToArray();
        text.AppendLine().AppendLine("## Observations requiring interpretation").AppendLine();
        if (observations.Length == 0) {
            text.AppendLine("No selected file produced warnings, errors, process failures, or timeouts in this run.").AppendLine();
        } else {
            text.AppendLine("| SHA-256 | Source | Format | Outcome | Diagnostics | Exception |");
            text.AppendLine("| --- | --- | --- | --- | --- | --- |");
            foreach (CorpusFileRecord file in observations) {
                string pdfFailures = file.PdfEvidence is null
                    ? string.Empty
                    : string.Join(", ", file.PdfEvidence.Stages.Where(static stage => !stage.Succeeded).Select(static stage => "pdf." + stage.Name));
                string diagnostics = string.Join(", ", file.DiagnosticCodes.Concat(string.IsNullOrEmpty(pdfFailures) ? Array.Empty<string>() : new[] { pdfFailures }));
                text.Append("| `").Append(ShortHash(file.Sha256)).Append("` | ").Append(E(file.SourceName ?? "withheld"))
                    .Append(" | ").Append(file.ContentKind).Append(" | ").Append(E(file.Outcome))
                    .Append(" | ").Append(E(diagnostics))
                    .Append(" | ").Append(E(file.ExceptionType ?? string.Empty)).AppendLine(" |");
            }
            text.AppendLine();
        }
        text.AppendLine("## Interpretation boundary").AppendLine();
        int sampleTarget = report.Configuration.MaxPerFormat;
        string upperBound = Math.Min(100d, 300d / sampleTarget)
            .ToString("0.##", CultureInfo.InvariantCulture);
        text.Append("A target of ").Append(sampleTarget)
            .AppendLine(" unique files per format is large enough to expose many recurring defects while remaining practical for a monthly isolated-process run. " +
                        "The familiar rule-of-three would put an approximate 95% upper bound near " + upperBound +
                        "% after " + sampleTarget + " independent observations with zero failures, " +
                        "but these corpus files are not an independent random sample. That calculation explains the sample budget; it is not a reliability guarantee.")
            .AppendLine();
        text.AppendLine("The lane proves only that OfficeIMO content-detected and attempted its normalized read contract for the recorded hashes under the recorded limits and runtime. " +
                        "The managed first-page render is an operational probe and does not assess visual fidelity. The lane also does not prove editing round trips, semantic preservation of every feature, safety of opening files in another application, or support for unmeasured files. " +
                        "Actionable findings should be minimized into provenance-tracked fixtures and moved to the owning format test suite.");
        return text.ToString().Replace("\r\n", "\n");
    }

    private static string ShortHash(string? value) => string.IsNullOrEmpty(value) ? "unavailable" : value[..Math.Min(16, value.Length)];
    private static string E(string value) {
        var escaped = new StringBuilder(value.Length);
        foreach (Rune rune in value.EnumerateRunes()) {
            UnicodeCategory category = Rune.GetUnicodeCategory(rune);
            if (Rune.IsControl(rune) || category == UnicodeCategory.Format ||
                category == UnicodeCategory.LineSeparator ||
                category == UnicodeCategory.ParagraphSeparator) {
                escaped.Append(' ');
            } else if (Rune.IsPunctuation(rune) || Rune.IsSymbol(rune)) {
                escaped.Append("&#").Append(rune.Value).Append(';');
            } else {
                escaped.Append(rune.ToString());
            }
        }
        return escaped.ToString();
    }
}
