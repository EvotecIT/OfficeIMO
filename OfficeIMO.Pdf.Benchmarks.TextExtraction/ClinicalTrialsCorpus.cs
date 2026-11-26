using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf.Benchmarks.TextExtraction;

/// <summary>
/// A reading-quality corpus of trial protocols registered on ClinicalTrials.gov. The registry record of each trial
/// is the ground truth: its title, identifier, sponsor, primary outcomes, arms, interventions and eligibility
/// criteria all appear in the protocol PDF the sponsor uploaded. The repository keeps only the manifest (trial
/// identifiers, file names and SHA-256 hashes); the PDFs and registry records are downloaded to a local folder.
/// </summary>
internal static class ClinicalTrialsCorpus {
    private const string ApiBase = "https://clinicaltrials.gov/api/v2/studies";
    private const string DocumentBase = "https://cdn.clinicaltrials.gov/large-docs";
    private const string UserAgent = "OfficeIMO-quality-corpus";
    private const int ManifestVersion = 1;

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>
    /// Picks the sample: for each sponsor class, every registered trial with a protocol PDF is ranked by
    /// SHA-256(seed + NCT number) and the lowest-ranked trials whose protocol fits the size limit are kept. The
    /// ranking is a uniform random sample that is reproducible from the seed and does not depend on API order.
    /// </summary>
    internal static int Sample(string[] args) {
        string manifestPath = Path.GetFullPath(Options.Get(args, "--manifest") ?? throw new ArgumentException("--manifest is required."));
        string seed = Options.Get(args, "--seed") ?? "officeimo-ctgov-1";
        long maxBytes = long.Parse(Options.Get(args, "--max-file-bytes") ?? (25L * 1024 * 1024).ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        Dictionary<string, int> perClass = ParseCounts(Options.Get(args, "--per-class")
            ?? "INDUSTRY=50,OTHER=50,NIH=25,FED=25,OTHER_GOV=20,NETWORK=20,INDIV=10");

        using HttpClient http = CreateClient();
        var trials = new List<JsonObject>();
        foreach ((string sponsorClass, int count) in perClass) {
            var candidates = new List<(string Rank, JsonObject Trial)>();
            string? pageToken = null;
            do {
                string url = $"{ApiBase}?format=json&pageSize=1000&fields=NCTId,LeadSponsorClass,LastUpdatePostDate,LargeDocumentModule"
                    + $"&filter.advanced={Uri.EscapeDataString($"AREA[LargeDocHasProtocol]true AND AREA[LeadSponsorClass]{sponsorClass}")}"
                    + (pageToken is null ? string.Empty : "&pageToken=" + Uri.EscapeDataString(pageToken));
                JsonNode page = GetJson(http, url);
                foreach (JsonNode? study in page["studies"]?.AsArray() ?? []) {
                    JsonObject? trial = ToCandidate(study, sponsorClass, maxBytes);
                    if (trial is not null) {
                        candidates.Add((Rank(seed, (string)trial["nctId"]!), trial));
                    }
                }
                pageToken = (string?)page["nextPageToken"];
            } while (pageToken is not null);
            List<JsonObject> chosen = candidates.OrderBy(static c => c.Rank, StringComparer.Ordinal).Take(count).Select(static c => c.Trial).ToList();
            Console.WriteLine($"{sponsorClass}: {candidates.Count} trials with a protocol PDF of at most {maxBytes / (1024 * 1024)} MB, kept {chosen.Count}.");
            trials.AddRange(chosen);
        }

        var manifest = new JsonObject {
            ["version"] = ManifestVersion,
            ["source"] = "ClinicalTrials.gov API v2, trials whose registry record includes a protocol document",
            ["seed"] = seed,
            ["maxFileBytes"] = maxBytes,
            ["perClass"] = new JsonObject(perClass.Select(static p => KeyValuePair.Create(p.Key, (JsonNode?)p.Value))),
            ["trials"] = new JsonArray(trials.OrderBy(static t => (string)t["sponsorClass"]!, StringComparer.Ordinal)
                .ThenBy(static t => (string)t["nctId"]!, StringComparer.Ordinal).Select(static t => (JsonNode)t).ToArray())
        };
        Directory.CreateDirectory(Path.GetDirectoryName(manifestPath)!);
        File.WriteAllText(manifestPath, manifest.ToJsonString(JsonOptions) + "\n");
        Console.WriteLine($"Wrote {trials.Count} trials to {manifestPath}. Run ctgov-fetch --record-hashes to download and pin them.");
        return 0;
    }

    /// <summary>
    /// Downloads each protocol PDF to <c>&lt;root&gt;/pdfs/&lt;NCT&gt;.pdf</c> and the current registry record to
    /// <c>&lt;root&gt;/records/&lt;NCT&gt;.json</c>. Files already present with the pinned hash are kept. A file
    /// whose hash differs from the manifest is reported and left out of <c>pdfs</c>, so scores always cover the
    /// pinned bytes. <c>--record-hashes</c> writes the hashes of a first download into the manifest.
    /// </summary>
    internal static int Fetch(string[] args) {
        string manifestPath = Path.GetFullPath(Options.Get(args, "--manifest") ?? throw new ArgumentException("--manifest is required."));
        string root = Path.GetFullPath(Options.Get(args, "--root") ?? throw new ArgumentException("--root is required."));
        bool recordHashes = Options.Has(args, "--record-hashes");
        JsonObject manifest = LoadManifest(manifestPath);
        string pdfs = Directory.CreateDirectory(Path.Combine(root, "pdfs")).FullName;
        string records = Directory.CreateDirectory(Path.Combine(root, "records")).FullName;

        using HttpClient http = CreateClient();
        int downloaded = 0, kept = 0, mismatched = 0, failed = 0, changedRecords = 0;
        bool hashesAssigned = false;
        foreach (JsonObject trial in Trials(manifest)) {
            string nctId = (string)trial["nctId"]!;
            string pdfPath = Path.Combine(pdfs, nctId + ".pdf");
            string? pinned = (string?)trial["sha256"];
            try {
                if (!(File.Exists(pdfPath) && pinned is not null && Sha256(File.ReadAllBytes(pdfPath)) == pinned)) {
                    // A cached file that doesn't match must not survive a failed download and be scored.
                    File.Delete(pdfPath);
                    byte[] bytes = GetBytes(http, $"{DocumentBase}/{nctId[^2..]}/{nctId}/{(string)trial["file"]!}");
                    string hash = Sha256(bytes);
                    if (pinned is null || recordHashes) {
                        hashesAssigned |= hash != pinned;
                        trial["sha256"] = hash;
                    } else if (hash != pinned) {
                        Console.Error.WriteLine($"{nctId}: the protocol no longer matches the pinned hash; left out.");
                        File.Delete(pdfPath);
                        mismatched++;
                        continue;
                    }
                    File.WriteAllBytes(pdfPath, bytes);
                    downloaded++;
                } else {
                    kept++;
                }

                JsonNode record = GetJson(http, $"{ApiBase}/{nctId}?format=json");
                string? updated = (string?)record["protocolSection"]?["statusModule"]?["lastUpdatePostDateStruct"]?["date"];
                if (updated is not null && updated != (string?)trial["lastUpdatePosted"]) {
                    changedRecords++;
                    Console.Error.WriteLine($"{nctId}: registry record updated on {updated} (manifest has {(string?)trial["lastUpdatePosted"]}).");
                }
                File.WriteAllText(Path.Combine(records, nctId + ".json"), record.ToJsonString(JsonOptions));
            } catch (Exception exception) when (exception is HttpRequestException or TaskCanceledException) {
                Console.Error.WriteLine($"{nctId}: {exception.Message}");
                failed++;
            }
        }
        if (hashesAssigned) {
            File.WriteAllText(manifestPath, manifest.ToJsonString(JsonOptions) + "\n");
        }
        Console.WriteLine($"Downloaded {downloaded}, already present {kept}, hash mismatches {mismatched}, failures {failed}, records updated since sampling {changedRecords}.");
        return failed == 0 && mismatched == 0 ? 0 : 1;
    }

    /// <summary>
    /// Scores the texts a <c>run --write-text</c> stored for each engine against the registry records and writes
    /// <c>ctgov-fields.csv</c> (one row per field and engine) and <c>ctgov-summary.md</c>.
    /// </summary>
    internal static int Score(string[] args) {
        string manifestPath = Path.GetFullPath(Options.Get(args, "--manifest") ?? throw new ArgumentException("--manifest is required."));
        string root = Path.GetFullPath(Options.Get(args, "--root") ?? throw new ArgumentException("--root is required."));
        string output = Path.GetFullPath(Options.Get(args, "--output") ?? throw new ArgumentException("--output is required."));
        string suite = Options.Get(args, "--suite") ?? "ctgov";
        JsonObject manifest = LoadManifest(manifestPath);
        Dictionary<(string Engine, string File), (string Status, double Ms)> timings = ReadTimings(output, suite);
        string[] engines = timings.Keys.Select(static k => k.Engine).Distinct().OrderBy(static e => e, StringComparer.Ordinal).ToArray();
        if (engines.Length == 0) {
            throw new InvalidOperationException($"No results for suite '{suite}' were found in {output}. Run the corpus with run --suite {suite}=<root>/pdfs --write-text first.");
        }

        var scores = new List<FieldScore>();
        var documents = new List<DocumentScore>();
        foreach (JsonObject trial in Trials(manifest)) {
            string nctId = (string)trial["nctId"]!;
            string sponsorClass = (string)trial["sponsorClass"]!;
            string recordPath = Path.Combine(root, "records", nctId + ".json");
            string pdfPath = Path.Combine(root, "pdfs", nctId + ".pdf");
            if (!File.Exists(recordPath) || !File.Exists(pdfPath)) {
                Console.Error.WriteLine($"{nctId}: not downloaded; skipped.");
                continue;
            }
            List<RegistryField> fields = RegistryField.From(JsonNode.Parse(File.ReadAllText(recordPath))!);
            int pages = CountPages(pdfPath);
            foreach (string engine in engines) {
                if (!timings.TryGetValue((engine, nctId + ".pdf"), out var timing)) {
                    continue;
                }
                string textPath = Path.Combine(output, "texts", engine, suite, nctId + ".pdf.txt");
                bool ok = timing.Status == "ok";
                if (ok && !File.Exists(textPath)) {
                    throw new FileNotFoundException($"{engine} read {nctId} but its text is missing. Run the corpus with --write-text, or run it again for this file.", textPath);
                }
                documents.Add(new DocumentScore(nctId, sponsorClass, engine, timing.Status, pages, timing.Ms));
                ProtocolText? text = ok ? new ProtocolText(File.ReadAllText(textPath)) : null;
                foreach (RegistryField field in fields) {
                    scores.Add(text is null
                        ? new FieldScore(nctId, sponsorClass, engine, field, Match.Missing, false, false, 0)
                        : text.Score(nctId, sponsorClass, engine, field));
                }
            }
        }

        WriteFieldCsv(Path.Combine(output, "ctgov-fields.csv"), scores);
        string summary = Summarize(scores, documents, engines);
        File.WriteAllText(Path.Combine(output, "ctgov-summary.md"), summary);
        Console.WriteLine(summary);
        return 0;
    }

    private static JsonObject? ToCandidate(JsonNode? study, string sponsorClass, long maxBytes) {
        string? nctId = (string?)study?["protocolSection"]?["identificationModule"]?["nctId"];
        JsonArray? docs = study?["documentSection"]?["largeDocumentModule"]?["largeDocs"]?.AsArray();
        if (nctId is null || docs is null) {
            return null;
        }
        // The first document that contains the protocol; combined protocol, SAP and consent uploads count.
        JsonNode? protocol = docs.FirstOrDefault(static d => (bool?)d?["hasProtocol"] == true
            && ((string?)d?["filename"])?.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) == true);
        long size = (long?)protocol?["size"] ?? 0;
        if (protocol is null || size <= 0 || size > maxBytes) {
            return null;
        }
        return new JsonObject {
            ["nctId"] = nctId,
            ["sponsorClass"] = sponsorClass,
            ["lastUpdatePosted"] = (string?)study!["protocolSection"]?["statusModule"]?["lastUpdatePostDateStruct"]?["date"],
            ["file"] = (string)protocol["filename"]!,
            ["bytes"] = size,
            ["sha256"] = null
        };
    }

    private static string Summarize(List<FieldScore> scores, List<DocumentScore> documents, string[] engines) {
        var report = new StringBuilder();
        report.AppendLine("# ClinicalTrials.gov protocol reading");
        report.AppendLine();
        report.AppendLine($"{documents.Select(static d => d.NctId).Distinct().Count()} protocols, {scores.Select(static s => (s.NctId, s.Field.Category, s.Field.Index)).Distinct().Count()} registry fields.");
        report.AppendLine("Intact: the field's words appear contiguously. In order: they all appear in order, possibly broken up. Words: at least 90% of its words appear somewhere.");
        report.AppendLine("Heading: a line that names the field's section exists. In section: an occurrence of the field follows, within 2,000 words, a line that opens with the field's section name, with no other section's opening line in between.");
        report.AppendLine();

        string[] classes = documents.Select(static d => d.SponsorClass).Distinct().OrderBy(static c => c, StringComparer.Ordinal).Append("all").ToArray();
        report.AppendLine("## By sponsor class");
        report.AppendLine();
        report.AppendLine("| Sponsor | Engine | Files | Read | Intact | In order | Words | Heading | In section | ms/page p50 | ms/page p90 |");
        report.AppendLine("|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|");
        foreach (string sponsorClass in classes) {
            foreach (string engine in engines) {
                List<FieldScore> s = scores.Where(x => x.Engine == engine && (sponsorClass == "all" || x.SponsorClass == sponsorClass)).ToList();
                List<DocumentScore> d = documents.Where(x => x.Engine == engine && (sponsorClass == "all" || x.SponsorClass == sponsorClass)).ToList();
                double[] perPage = d.Where(static x => x.Status == "ok" && x.Pages > 0).Select(static x => x.Ms / x.Pages).OrderBy(static x => x).ToArray();
                List<FieldScore> sectioned = s.Where(static x => x.Field.Section is not null).ToList();
                report.AppendLine($"| {sponsorClass} | {engine} | {d.Count} | {d.Count(static x => x.Status == "ok")} | {Percent(s, static x => x.Match == Match.Intact)} | {Percent(s, static x => x.Match <= Match.InOrder)} | {Percent(s, static x => x.Match <= Match.Words)} | {Percent(sectioned, static x => x.HeadingFound)} | {Percent(sectioned, static x => x.InSection)} | {Quantile(perPage, 0.5)} | {Quantile(perPage, 0.9)} |");
            }
        }

        report.AppendLine();
        report.AppendLine("## By field");
        report.AppendLine();
        report.AppendLine("| Field | Engine | Fields | Intact | In order | Words | Heading | In section |");
        report.AppendLine("|---|---|---:|---:|---:|---:|---:|---:|");
        foreach (string category in RegistryField.Categories) {
            foreach (string engine in engines) {
                List<FieldScore> s = scores.Where(x => x.Engine == engine && x.Field.Category == category).ToList();
                bool sectioned = s.Any(static x => x.Field.Section is not null);
                report.AppendLine($"| {category} | {engine} | {s.Count} | {Percent(s, static x => x.Match == Match.Intact)} | {Percent(s, static x => x.Match <= Match.InOrder)} | {Percent(s, static x => x.Match <= Match.Words)} | {(sectioned ? Percent(s, static x => x.HeadingFound) : "–")} | {(sectioned ? Percent(s, static x => x.InSection) : "–")} |");
            }
        }

        report.AppendLine();
        report.AppendLine("## Protocols with the lowest intact share");
        report.AppendLine();
        report.AppendLine("| Trial | Sponsor | Pages | " + string.Join(" | ", engines.Select(static e => e + " intact")) + " | Most common outcome |");
        report.AppendLine("|---|---|---:|" + string.Concat(engines.Select(static _ => "---:|")) + "---|");
        string first = engines.Contains("officeimo-read") ? "officeimo-read" : engines[0];
        var worst = scores.Where(x => x.Engine == first).GroupBy(static x => x.NctId)
            .Select(g => (NctId: g.Key, Share: g.Count(static x => x.Match == Match.Intact) / (double)g.Count(), Scores: g.ToList()))
            .OrderBy(static g => g.Share).ThenBy(static g => g.NctId, StringComparer.Ordinal).Take(15);
        foreach (var trial in worst) {
            DocumentScore doc = documents.First(d => d.NctId == trial.NctId && d.Engine == first);
            IEnumerable<string> shares = engines.Select(e => Percent(scores.Where(x => x.Engine == e && x.NctId == trial.NctId).ToList(), static x => x.Match == Match.Intact));
            string common = doc.Status != "ok" ? doc.Status
                : trial.Scores.Where(static x => x.Match != Match.Intact).GroupBy(static x => x.Field.Category + " " + x.Match)
                    .OrderByDescending(static g => g.Count()).Select(static g => $"{g.Key} ({g.Count()})").FirstOrDefault() ?? "–";
            report.AppendLine($"| {trial.NctId} | {doc.SponsorClass} | {doc.Pages} | {string.Join(" | ", shares)} | {common} |");
        }
        return report.ToString();
    }

    private static string Percent<T>(List<T> items, Func<T, bool> predicate) =>
        items.Count == 0 ? "–" : (100.0 * items.Count(predicate) / items.Count).ToString("F1", CultureInfo.InvariantCulture) + "%";

    private static string Quantile(double[] sorted, double fraction) =>
        sorted.Length == 0 ? "–" : sorted[Math.Min(sorted.Length - 1, (int)Math.Ceiling(fraction * sorted.Length) - 1)].ToString("F1", CultureInfo.InvariantCulture);

    private static void WriteFieldCsv(string path, List<FieldScore> scores) {
        using var csv = new StreamWriter(path, append: false, new UTF8Encoding(false));
        csv.WriteLine("trial,sponsor_class,engine,field,index,words,match,heading_found,in_section,word_share");
        foreach (FieldScore s in scores) {
            csv.WriteLine(string.Join(',', s.NctId, s.SponsorClass, s.Engine, s.Field.Category, s.Field.Index.ToString(CultureInfo.InvariantCulture),
                s.Field.Tokens.Length.ToString(CultureInfo.InvariantCulture), s.Match.ToString(), s.HeadingFound ? "1" : "0", s.InSection ? "1" : "0",
                s.WordShare.ToString("F3", CultureInfo.InvariantCulture)));
        }
    }

    private static Dictionary<(string Engine, string File), (string Status, double Ms)> ReadTimings(string output, string suite) {
        var timings = new Dictionary<(string, string), (string, double)>();
        foreach (string path in Directory.EnumerateFiles(output, "results-*.csv").OrderBy(static p => File.GetLastWriteTimeUtc(p))) {
            foreach (string[] fields in ResultSummary.ReadCsv(path).Skip(1)) {
                if (fields.Length >= 6 && fields[0] == suite
                    && double.TryParse(fields[5], NumberStyles.Float, CultureInfo.InvariantCulture, out double ms)) {
                    timings[(fields[2], fields[1])] = (fields[4], ms);
                }
            }
        }
        return timings;
    }

    /// <summary>Page counts come from PdfPig so that neither scored OfficeIMO path decides the per-page divisor.</summary>
    private static int CountPages(string path) {
        try {
            using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(path);
            return document.NumberOfPages;
        } catch (Exception) {
            return 0;
        }
    }

    private static JsonObject LoadManifest(string path) {
        JsonObject manifest = JsonNode.Parse(File.ReadAllText(path))?.AsObject() ?? throw new InvalidDataException("Empty manifest.");
        if ((int?)manifest["version"] != ManifestVersion) {
            throw new InvalidDataException($"Manifest version {(int?)manifest["version"]} is not supported.");
        }
        return manifest;
    }

    private static IEnumerable<JsonObject> Trials(JsonObject manifest) =>
        (manifest["trials"]?.AsArray() ?? []).Select(static t => t!.AsObject());

    private static Dictionary<string, int> ParseCounts(string value) =>
        value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(static part => part.Split('=', 2))
            .ToDictionary(static p => p[0], static p => int.Parse(p[1], CultureInfo.InvariantCulture), StringComparer.Ordinal);

    private static string Rank(string seed, string nctId) => Sha256(Encoding.UTF8.GetBytes(seed + ":" + nctId));

    private static string Sha256(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

    private static HttpClient CreateClient() {
        var http = new HttpClient { Timeout = TimeSpan.FromMinutes(2) };
        http.DefaultRequestHeaders.UserAgent.ParseAdd(UserAgent);
        return http;
    }

    private static JsonNode GetJson(HttpClient http, string url) =>
        JsonNode.Parse(Encoding.UTF8.GetString(GetBytes(http, url))) ?? throw new InvalidDataException("Empty response from " + url);

    /// <summary>One request at a time with a short pause, and a few retries on throttling or server errors.</summary>
    private static byte[] GetBytes(HttpClient http, string url) {
        for (int attempt = 1; ; attempt++) {
            Thread.Sleep(250);
            using HttpResponseMessage response = http.GetAsync(url).GetAwaiter().GetResult();
            bool retry = response.StatusCode == HttpStatusCode.TooManyRequests || (int)response.StatusCode >= 500;
            if (retry && attempt < 4) {
                Thread.Sleep(TimeSpan.FromSeconds(5 * attempt));
                continue;
            }
            response.EnsureSuccessStatusCode();
            return response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
        }
    }
}

internal enum Match {
    Intact,
    InOrder,
    Words,
    Missing
}

internal sealed record DocumentScore(string NctId, string SponsorClass, string Engine, string Status, int Pages, double Ms);

internal sealed record FieldScore(string NctId, string SponsorClass, string Engine, RegistryField Field, Match Match, bool HeadingFound, bool InSection, double WordShare);

/// <summary>One registry value expected in the protocol, with the section heading words it belongs under.</summary>
internal sealed record RegistryField(string Category, int Index, string[] Tokens, string[]? Section) {
    internal static readonly string[] Categories = ["nct-id", "title", "sponsor", "primary-outcome", "arm", "intervention", "inclusion", "exclusion"];

    private static readonly string[] OutcomeSection = ["endpoint", "endpoints", "outcome", "outcomes"];
    private static readonly string[] ArmSection = ["arm", "arms", "design", "group", "groups", "cohort", "cohorts", "randomization", "treatment", "treatments"];
    private static readonly string[] InterventionSection = ["intervention", "interventions", "drug", "drugs", "treatment", "treatments", "product", "products", "device", "therapy"];
    private static readonly string[] InclusionSection = ["inclusion", "eligibility", "selection", "population"];
    private static readonly string[] ExclusionSection = ["exclusion", "eligibility", "selection", "population"];

    /// <summary>Every section name a field can belong to; a line opening with one of them starts a section.</summary>
    internal static readonly HashSet<string> SectionWords = new(
        OutcomeSection.Concat(ArmSection).Concat(InterventionSection).Concat(InclusionSection).Concat(ExclusionSection), StringComparer.Ordinal);

    /// <summary>
    /// Reads the fields a protocol should contain. Eligibility criteria are split into one field per criterion
    /// with at least five words, so long criteria test line and column order.
    /// </summary>
    internal static List<RegistryField> From(JsonNode record) {
        JsonNode? protocol = record["protocolSection"];
        var fields = new List<RegistryField>();
        void Add(string category, string? value, string[]? section, int minWords = 1) {
            string[] tokens = ProtocolText.Tokenize(value ?? string.Empty);
            if (tokens.Length >= minWords) {
                fields.Add(new RegistryField(category, fields.Count(f => f.Category == category), tokens, section));
            }
        }

        Add("nct-id", (string?)protocol?["identificationModule"]?["nctId"], null);
        Add("title", (string?)protocol?["identificationModule"]?["officialTitle"] ?? (string?)protocol?["identificationModule"]?["briefTitle"], null);
        Add("sponsor", (string?)protocol?["sponsorCollaboratorsModule"]?["leadSponsor"]?["name"], null);
        foreach (JsonNode? outcome in protocol?["outcomesModule"]?["primaryOutcomes"]?.AsArray() ?? []) {
            Add("primary-outcome", (string?)outcome?["measure"], OutcomeSection);
        }
        foreach (JsonNode? arm in protocol?["armsInterventionsModule"]?["armGroups"]?.AsArray() ?? []) {
            Add("arm", (string?)arm?["label"], ArmSection);
        }
        foreach (JsonNode? intervention in protocol?["armsInterventionsModule"]?["interventions"]?.AsArray() ?? []) {
            Add("intervention", (string?)intervention?["name"], InterventionSection);
        }
        string category = "inclusion";
        foreach (string line in ((string?)protocol?["eligibilityModule"]?["eligibilityCriteria"] ?? string.Empty).Split('\n')) {
            // The registry writes the subheadings as "Inclusion Criteria:" or "* EXCLUSION CRITERIA:".
            string trimmed = StripListMarker(line.Trim()).Trim();
            if (trimmed.StartsWith("exclusion", StringComparison.OrdinalIgnoreCase) && trimmed.Length < 40) {
                category = "exclusion";
                continue;
            }
            if (trimmed.StartsWith("inclusion", StringComparison.OrdinalIgnoreCase) && trimmed.Length < 40) {
                category = "inclusion";
                continue;
            }
            Add(category, trimmed, category == "inclusion" ? InclusionSection : ExclusionSection, minWords: 5);
        }
        return fields;
    }

    /// <summary>Removes bullets and list numbering such as <c>*</c>, <c>-</c>, <c>1.</c>, <c>(a)</c> or <c>ii)</c>.</summary>
    private static string StripListMarker(string line) => ListMarker.Replace(line, string.Empty);

    private static readonly Regex ListMarker = new(@"^(?:\s*(?:[*\-•·▪●○]|\(?(?:\d{1,3}|[a-zA-Z]|[ivxIVX]{1,4})[.)])\s+)+", RegexOptions.CultureInvariant);
}

/// <summary>
/// An engine's extracted text, normalised for matching: NFKC, lower case, hyphens removed (a word broken across
/// lines with a hyphen joins up again) and every other non-letter, non-digit character treated as a separator.
/// </summary>
internal sealed class ProtocolText {
    private const int MaxStarts = 4000;
    private readonly string[] _tokens;
    private readonly Dictionary<string, List<int>> _positions = new(StringComparer.Ordinal);
    private readonly HashSet<string> _vocabulary;
    private const int SectionReach = 2000;
    private readonly List<string[]> _headings = [];
    private readonly List<(int Token, int HeadingEnd, string[] Opening)> _sectionStarts = [];
    private readonly int _frontMatterEnd;

    internal ProtocolText(string text) {
        var tokens = new List<string>();
        foreach (string line in JoinHyphenatedLines(text).Split('\n')) {
            string[] lineTokens = Tokenize(line);
            if (lineTokens.Length is > 0 and <= 12 && LooksLikeHeading(line)) {
                _headings.Add(lineTokens);
            }
            if (OpensSection(line, out int headingWords) is { } opening) {
                // A field matched inside the heading itself ("Treatment" as an arm label) doesn't follow it.
                _sectionStarts.Add((tokens.Count, tokens.Count + headingWords, opening));
            }
            tokens.AddRange(lineTokens);
        }
        _tokens = tokens.ToArray();
        for (int i = 0; i < _tokens.Length; i++) {
            if (!_positions.TryGetValue(_tokens[i], out List<int>? list)) {
                _positions[_tokens[i]] = list = [];
            }
            list.Add(i);
        }
        _vocabulary = [.. _positions.Keys];
        _frontMatterEnd = Math.Max(2000, _tokens.Length / 20);
    }

    internal FieldScore Score(string nctId, string sponsorClass, string engine, RegistryField field) {
        string[] f = field.Tokens;
        double wordShare = f.Count(_vocabulary.Contains) / (double)f.Length;
        (Match match, List<int> at) = Find(f);
        if (match == Match.Missing && wordShare >= 0.9) {
            match = Match.Words;
        }
        bool headingFound;
        bool inSection;
        if (field.Section is null) {
            // Identifier, title and sponsor belong on the title page or in the synopsis at the front.
            headingFound = false;
            inSection = match <= Match.InOrder && at.Any(position => position < _frontMatterEnd);
        } else {
            headingFound = _headings.Any(h => h.Any(field.Section.Contains));
            inSection = match <= Match.InOrder && at.Any(position => NearestSectionStart(position) is { } opening && opening.Any(field.Section.Contains));
        }
        return new FieldScore(nctId, sponsorClass, engine, field, match, headingFound, inSection, wordShare);
    }

    /// <summary>
    /// The best placement of the field and where it occurs: every contiguous run of its words, otherwise every place
    /// all of its words appear in order within a window of three times its length. The search starts at each occurrence
    /// of the field's rarest word, so a field that appears in the synopsis and again in its section is found in both.
    /// </summary>
    private (Match Match, List<int> At) Find(string[] f) {
        // At lists every qualifying occurrence, intact or in order, so a field intact in the synopsis and broken up
        // under its section still counts as placed in its section.
        int anchor = 0;
        int anchorCount = int.MaxValue;
        for (int i = 0; i < f.Length; i++) {
            int count = _positions.TryGetValue(f[i], out List<int>? p) ? p.Count : 0;
            if (count == 0) {
                return (Match.Missing, []);
            }
            if (count < anchorCount) {
                anchor = i;
                anchorCount = count;
            }
        }
        List<int> starts = _positions[f[anchor]];
        var intact = new List<int>();
        var inOrder = new List<int>();
        for (int s = 0; s < Math.Min(starts.Count, MaxStarts); s++) {
            int start = starts[s] - anchor;
            if (start >= 0 && start + f.Length <= _tokens.Length && Contiguous(f, start)) {
                intact.Add(start);
            } else if (InOrder(f, anchor, starts[s]) is int first and >= 0) {
                inOrder.Add(first);
            }
        }
        if (intact.Count == 0) {
            return inOrder.Count > 0 ? (Match.InOrder, inOrder) : (Match.Missing, []);
        }
        intact.AddRange(inOrder);
        return (Match.Intact, intact);
    }

    private bool Contiguous(string[] f, int start) {
        for (int i = 0; i < f.Length; i++) {
            if (!string.Equals(_tokens[start + i], f[i], StringComparison.Ordinal)) {
                return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Places the field's other words around one occurrence of its anchor word, each as close as possible, and
    /// returns where the placement starts, or -1 when the words from first to last span more than three times the
    /// field's length.
    /// </summary>
    private int InOrder(string[] f, int anchor, int anchorPosition) {
        int window = 3 * f.Length;
        int first = anchorPosition;
        for (int i = anchor - 1; i >= 0; i--) {
            first = Previous(f[i], first, anchorPosition - window + 1);
            if (first < 0) {
                return -1;
            }
        }
        int last = anchorPosition;
        for (int i = anchor + 1; i < f.Length; i++) {
            last = Next(f[i], last, first + window - 1);
            if (last < 0) {
                return -1;
            }
        }
        return first;
    }

    private int Next(string token, int after, int limit) {
        List<int> p = _positions[token];
        int index = p.BinarySearch(after + 1);
        index = index < 0 ? ~index : index;
        return index < p.Count && p[index] <= limit ? p[index] : -1;
    }

    private int Previous(string token, int before, int limit) {
        List<int> p = _positions[token];
        int index = p.BinarySearch(before - 1);
        index = index < 0 ? ~index - 1 : index;
        return index >= 0 && p[index] >= limit ? p[index] : -1;
    }

    /// <summary>
    /// The opening words of the nearest line before the field, within <see cref="SectionReach"/> words, that opens
    /// with a section name once its numbering is removed. Table-of-contents lines with dot leaders are ignored.
    /// </summary>
    private string[]? NearestSectionStart(int at) {
        string[]? nearest = null;
        foreach ((int token, int headingEnd, string[] opening) in _sectionStarts) {
            if (token > at) {
                break;
            }
            nearest = at >= headingEnd && at - token <= SectionReach ? opening : null;
        }
        return nearest;
    }

    /// <summary>
    /// The opening words of a line that starts a section: once numbering is removed, it starts with a section name,
    /// or with a qualifier such as "study" or "primary" followed by one ("Study drug", "Primary endpoints"). The line
    /// must be heading-length, or carry a colon early ("Inclusion criteria (all must be met): ..."), so prose that
    /// happens to start with "Treatment" doesn't open a section. Dot-leader table-of-contents lines are ignored.
    /// </summary>
    private static string[]? OpensSection(string line, out int headingWords) {
        headingWords = 0;
        if (line.Contains("....", StringComparison.Ordinal)) {
            return null;
        }
        string body = SectionNumber.Replace(line, string.Empty);
        // "Inclusion / Exclusion", "Inclusion & Exclusion" and "Inclusion and/or Exclusion" join names like "and" does.
        string[] words = Tokenize(Regex.Replace(body, @"\s*(?:and\s*/\s*or|/|&)\s*", " and ", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
        if (words.Length == 0) {
            return null;
        }
        int nameAt = RegistryField.SectionWords.Contains(words[0]) ? 0
            : words.Length > 1 && SectionQualifiers.Contains(words[0]) && RegistryField.SectionWords.Contains(words[1]) ? 1
            : -1;
        if (nameAt < 0) {
            return null;
        }
        int colon = body.IndexOf(':');
        int lineWords = Tokenize(line).Length;
        // Lengths are counted on the words as written, before a joiner is expanded. Without an early colon, a line
        // must also look like a heading, so a short sentence such as "Treatment was well tolerated." doesn't open a
        // section, while a title-case "Primary Outcomes." does.
        bool headingLength;
        if (colon >= 0 && Tokenize(body[..colon]).Length <= 8) {
            headingLength = true;
            headingWords = lineWords - Tokenize(body[(colon + 1)..]).Length;
        } else {
            headingLength = Tokenize(body).Length <= 8 && (LooksLikeHeading(body) || IsTitleCase(body));
            headingWords = lineWords;
        }
        // A joined heading ("Inclusion and exclusion criteria") opens both sections.
        int last = nameAt + 2 < words.Length && words[nameAt + 1] is "and" or "or" && RegistryField.SectionWords.Contains(words[nameAt + 2])
            ? nameAt + 2
            : nameAt;
        return headingLength ? words.Take(last + 1).ToArray() : null;
    }

    private static readonly HashSet<string> SectionQualifiers = new(StringComparer.Ordinal) {
        "study", "primary", "secondary", "efficacy", "safety", "investigational", "subject", "subjects", "patient", "patients",
        "participant", "participants", "trial", "clinical", "planned", "experimental"
    };

    /// <summary>A short line that does not end like a sentence. Table-of-contents lines count.</summary>
    /// <summary>Every word of four or more letters starts with a capital, as in "Primary Outcome Measures."</summary>
    private static bool IsTitleCase(string text) {
        string[] words = text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        return words.Any(static w => w.Length >= 4) && words.Where(static w => w.Length >= 4 && char.IsLetter(w[0])).All(static w => char.IsUpper(w[0]));
    }

    private static bool LooksLikeHeading(string line) {
        string trimmed = line.Trim();
        return trimmed.Length > 0 && !(trimmed.EndsWith('.') && !trimmed.EndsWith("..", StringComparison.Ordinal)) && !trimmed.EndsWith(',') && !trimmed.EndsWith(';');
    }

    /// <summary>Section numbering such as <c>5</c>, <c>5.2.1.</c>, <c>A.</c>, <c>IV.</c> or <c>(b)</c> at the start of a line.</summary>
    private static readonly Regex SectionNumber = new(@"^\s*(?:(?:section|part|chapter)\s+)?(?:\d+(?:\.\d+)*\.?|\(?[a-zA-Z][.)]|[ivxIVX]{1,5}[.)])\s+", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static string JoinHyphenatedLines(string text) {
        var builder = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            if (IsHyphen(c)) {
                int j = i + 1;
                while (j < text.Length && text[j] is ' ' or '\t' or '\r') {
                    j++;
                }
                if (j < text.Length && text[j] == '\n' && builder.Length > 0 && char.IsLetter(builder[^1])) {
                    j++;
                    while (j < text.Length && char.IsWhiteSpace(text[j])) {
                        j++;
                    }
                    i = j - 1;
                }
                continue;
            }
            builder.Append(c);
        }
        return builder.ToString();
    }

    internal static string[] Tokenize(string value) {
        string normalized = value.Normalize(NormalizationForm.FormKC);
        var tokens = new List<string>();
        var token = new StringBuilder();
        foreach (char c in normalized) {
            if (IsHyphen(c)) {
                continue;
            }
            if (char.IsLetterOrDigit(c)) {
                token.Append(char.ToLowerInvariant(c));
            } else if (token.Length > 0) {
                tokens.Add(token.ToString());
                token.Clear();
            }
        }
        if (token.Length > 0) {
            tokens.Add(token.ToString());
        }
        return tokens.ToArray();
    }

    private static bool IsHyphen(char c) => c is '-' or '­' or '‐' or '‑' or '‒' or '–' or '—' or '−';
}
