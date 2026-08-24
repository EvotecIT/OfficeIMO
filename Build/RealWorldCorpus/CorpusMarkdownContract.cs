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
                ReaderPolicy = new CorpusReaderPolicyConfiguration {
                    DetectionMode = ReaderDetectionMode.PreferContent,
                    InspectContainers = true,
                    DetectionMaxProbeBytes = 31,
                    DetectionMaxContainerEntries = 37,
                    ReadMaxCharacters = 41,
                    ReadMaxTableRows = 43,
                    ComputeHashes = false
                },
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
                    SourceName = Hostile("sourcename") + "\u202Espoof",
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
        if (markdown.Contains('\u202E') ||
            !markdown.Contains("A target of 17 unique files per format", StringComparison.Ordinal) ||
            !markdown.Contains("near 17.65% after 17 independent observations", StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                "Markdown interpretation or Unicode format-control neutralization is inconsistent.");
        }
        if (!markdown.Contains("| Samples per format | Total samples |", StringComparison.Ordinal) ||
            !markdown.Contains("| 0 | 0 | 0 | 0 | 17 | 23 |", StringComparison.Ordinal) ||
            !markdown.Contains("| Detection mode | Inspect containers | Detection probe bytes |", StringComparison.Ordinal) ||
            !markdown.Contains("| PreferContent | yes | 31 | 37 | 41 | 43 | no |", StringComparison.Ordinal) ||
            !markdown.Contains("| Package bytes | Package parts |", StringComparison.Ordinal) ||
            !markdown.Contains("| 29 | 0 | 0 | 0 | 0 | 0 |", StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                "Markdown does not record the sampling, reader, and package limits.");
        }
        Console.WriteLine("Dynamic Markdown fields are inert.");
        return 0;
    }

    private static string Hostile(string field) => $"![{field}](target)";
    private static string Inert(string field) => $"&#33;&#91;{field}&#93;&#40;target&#41;";
}
