using OfficeIMO.Reader;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusMarkdownContract {
    public static int Run() {
        var report = new CorpusReport {
            MeasurementStatus = Hostile("measurement"),
            StartedUtc = DateTimeOffset.UnixEpoch,
            CompletedUtc = DateTimeOffset.UnixEpoch,
            Provenance = new CorpusProvenance {
                CorpusId = Hostile("corpus"),
                SourceUri = Hostile("source"),
                ArchiveSha256 = Hostile("archive")
            },
            Environment = new CorpusEnvironment {
                Framework = Hostile("framework"),
                OperatingSystem = Hostile("operatingsystem")
            },
            Configuration = new CorpusConfiguration {
                MaxPerFormat = 17,
                MaxTotal = 23,
                PackagePolicy = new CorpusPackagePolicyConfiguration {
                    MaxPackageBytes = 29
                }
            },
            Totals = new CorpusTotals { Discovered = 1, EligibleUnique = 1, Selected = 1, Failed = 1 },
            Strata = new[] {
                new CorpusStratum { Format = ReaderInputKind.Html, EligibleUnique = 1, RequestedMaximum = 1, Selected = 1, Failed = 1 }
            },
            Files = new[] {
                new CorpusFileRecord {
                    Sha256 = new string('a', 64),
                    SourceName = Hostile("sourcename"),
                    ContentKind = ReaderInputKind.Html,
                    Selected = true,
                    Outcome = Hostile("outcome"),
                    DiagnosticCodes = new[] { Hostile("diagnostic") },
                    ExceptionType = Hostile("exception")
                }
            }
        };

        string markdown = CorpusReportWriter.RenderMarkdown(report);
        string[] fields = {
            "measurement", "corpus", "source", "archive", "framework", "operatingsystem",
            "sourcename", "outcome", "diagnostic", "exception"
        };
        foreach (string field in fields) {
            if (markdown.Contains(Hostile(field), StringComparison.Ordinal) ||
                !markdown.Contains(Inert(field), StringComparison.Ordinal)) {
                throw new InvalidOperationException($"Dynamic Markdown field '{field}' is not inert.");
            }
        }
        if (!markdown.Contains("| Samples per format | Total samples | Package bytes |", StringComparison.Ordinal) ||
            !markdown.Contains("| 0 | 0 | 0 | 0 | 17 | 23 | 29 |", StringComparison.Ordinal)) {
            throw new InvalidOperationException("Markdown does not record the sampling and package limits.");
        }
        Console.WriteLine("Dynamic Markdown fields are inert.");
        return 0;
    }

    private static string Hostile(string field) => $"![{field}](target)";
    private static string Inert(string field) => $"&#33;&#91;{field}&#93;&#40;target&#41;";
}
