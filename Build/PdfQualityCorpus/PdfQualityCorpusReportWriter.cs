using System.Text;
using System.Text.Json;

namespace OfficeIMO.PdfQualityCorpus;

internal static class PdfQualityCorpusReportWriter {
    internal static void Write(QualityReport report, string jsonPath, string markdownPath) {
        WriteAtomic(jsonPath, JsonSerializer.Serialize(report, QualityJson.Options));
        WriteAtomic(markdownPath, RenderMarkdown(report));
    }

    internal static string RenderMarkdown(QualityReport report) {
        var text = new StringBuilder();
        text.AppendLine("# OfficeIMO PDF quality corpus scorecard");
        text.AppendLine();
        text.AppendLine("This report records measured behavior for provenance-bound PDF files. It is not a claim about unmeasured documents.");
        text.AppendLine();
        text.AppendLine("## Run");
        text.AppendLine();
        text.AppendLine("| Field | Value |");
        text.AppendLine("| --- | --- |");
        Row(text, "Authority", report.Authority);
        Row(text, "Started UTC", report.StartedUtc.ToString("O", System.Globalization.CultureInfo.InvariantCulture));
        Row(text, "Completed UTC", report.CompletedUtc.ToString("O", System.Globalization.CultureInfo.InvariantCulture));
        Row(text, "Framework", report.Environment.Framework);
        Row(text, "Operating system", report.Environment.OperatingSystem);
        Row(text, "Architecture", report.Environment.ProcessArchitecture);
        Row(text, "OfficeIMO.Pdf assembly", report.Environment.EngineAssemblyVersion);
        Row(text, "Manifest", report.Configuration.ManifestFileName);
        Row(text, "Manifest SHA256", report.Configuration.ManifestSha256);
        Row(text, "Cases", report.Totals.Cases.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Row(text, "Passed", report.Totals.Passed.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Row(text, "Failed", report.Totals.Failed.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Row(text, "Timed out", report.Totals.TimedOut.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Row(text, "Operational score", Percent(report.Totals.OperationalScore));
        Row(text, "Expectation score", Percent(report.Totals.ExpectationScore));
        Row(text, "Peak worker memory", MebiBytes(report.Totals.PeakWorkingSetBytes));
        Row(text, "Worker memory budget", MebiBytes(report.Configuration.MaxWorkerMemoryBytes));
        text.AppendLine();
        text.AppendLine("## Sources");
        text.AppendLine();
        text.AppendLine("| Id | Repository | Commit | License |");
        text.AppendLine("| --- | --- | --- | --- |");
        foreach (QualitySource source in report.Sources) {
            text.Append("| ").Append(E(source.Id)).Append(" | ").Append(E(source.Repository)).Append(" | `")
                .Append(E(source.Commit)).Append("` | ").Append(E(source.License)).AppendLine(" |");
        }
        text.AppendLine();
        text.AppendLine("## Cases");
        text.AppendLine();
        text.AppendLine("| Case | Outcome | Pages | Text | Fonts | Embedded | Images | Rendered | Operations | Expectations | Time ms | Peak MiB |");
        text.AppendLine("| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |");
        foreach (QualityCaseResult item in report.Cases) {
            text.Append("| ").Append(E(item.Id)).Append(" | ").Append(E(item.Outcome)).Append(" | ")
                .Append(item.Metrics.PageCount).Append(" | ").Append(item.Metrics.TextCharacters).Append(" | ")
                .Append(item.Metrics.FontCount).Append(" | ").Append(item.Metrics.EmbeddedFontCount).Append(" | ")
                .Append(item.Metrics.ImageCount).Append(" | ").Append(item.Metrics.RenderSucceededPages).Append('/').Append(item.Metrics.RenderAttemptedPages).Append(" | ")
                .Append(Percent(item.OperationalScore)).Append(" | ").Append(Percent(item.ExpectationScore)).Append(" | ")
                .Append(item.WorkerWallClockMilliseconds).Append(" | ")
                .Append((item.PeakWorkingSetBytes / (1024D * 1024D)).ToString("F2", System.Globalization.CultureInfo.InvariantCulture)).AppendLine(" |");
        }
        QualityCaseResult[] failures = report.Cases.Where(item => item.Outcome != "passed").ToArray();
        if (failures.Length > 0) {
            text.AppendLine();
            text.AppendLine("## Failures");
            foreach (QualityCaseResult item in failures) {
                text.AppendLine();
                text.Append("### ").AppendLine(E(item.Id));
                if (!string.IsNullOrWhiteSpace(item.FailureCode)) text.AppendLine(E(item.FailureCode));
                foreach (QualityCheckResult check in item.Checks.Where(check => !check.Succeeded)) {
                    text.Append("1. Operation `").Append(E(check.Name)).Append("`: ").AppendLine(E(check.Message ?? check.ExceptionType ?? "failed"));
                }
                foreach (QualityExpectationResult expectation in item.Expectations.Where(expectation => !expectation.Succeeded)) {
                    text.Append("1. Expectation `").Append(E(expectation.Name)).Append("`: expected `")
                        .Append(E(expectation.Expected)).Append("`, actual `").Append(E(expectation.Actual)).AppendLine("`");
                }
            }
        }
        text.AppendLine();
        text.AppendLine("The operational score measures completed public API stages. The expectation score compares observed output with the pinned manifest. Feature counts are observations unless the manifest defines a minimum or maximum.");
        return text.ToString();
    }

    private static void Row(StringBuilder text, string name, string value) =>
        text.Append("| ").Append(E(name)).Append(" | ").Append(E(value)).AppendLine(" |");

    private static string Percent(double value) => value.ToString("P2", System.Globalization.CultureInfo.InvariantCulture);

    private static string MebiBytes(long bytes) =>
        (bytes / (1024D * 1024D)).ToString("F2", System.Globalization.CultureInfo.InvariantCulture) + " MiB";

    internal static string EscapeMarkdown(string value) {
        var text = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (character == '\r' || character == '\n' || character == '\t') {
                text.Append(' ');
            } else if ((character >= '\u202A' && character <= '\u202E') ||
                       (character >= '\u2066' && character <= '\u2069') ||
                       (char.IsControl(character) && character != '\t')) {
                text.Append('\uFFFD');
            } else {
                text.Append(character switch {
                    '&' => "&amp;",
                    '<' => "&lt;",
                    '>' => "&gt;",
                    '`' => "&#96;",
                    '\\' => "&#92;",
                    '|' => "&#124;",
                    '[' => "&#91;",
                    ']' => "&#93;",
                    '(' => "&#40;",
                    ')' => "&#41;",
                    '*' => "&#42;",
                    '_' => "&#95;",
                    '#' => "&#35;",
                    '+' => "&#43;",
                    '~' => "&#126;",
                    '!' => "&#33;",
                    _ => character.ToString()
                });
            }
        }
        return text.ToString();
    }

    private static string E(string value) => EscapeMarkdown(value);

    private static void WriteAtomic(string path, string content) {
        string fullPath = Path.GetFullPath(path);
        string directory = Path.GetDirectoryName(fullPath) ?? throw new InvalidOperationException("Report path has no directory.");
        Directory.CreateDirectory(directory);
        string temporary = Path.Combine(directory, "." + Path.GetFileName(fullPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        File.WriteAllText(temporary, content, new UTF8Encoding(false));
        File.Move(temporary, fullPath, overwrite: true);
    }
}
