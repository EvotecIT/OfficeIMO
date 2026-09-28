using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Html;
using OfficeIMO.Mhtml;

namespace OfficeIMO.Html.EditableEvidence;

internal static class EditableEvidenceRunner {
    private static readonly string[] SupportedTargets = ["word", "excel", "powerpoint", "onenote", "rtf", "markdown"];
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    internal static async Task<int> RunAsync(string[] args) {
        if (!TryParse(args, out Dictionary<string, string> options) ||
            !options.TryGetValue("case", out string? caseId) ||
            !options.TryGetValue("mhtml", out string? archivePath) ||
            !options.TryGetValue("output", out string? outputPath)) {
            Console.Error.WriteLine("Usage: HtmlEditableEvidence --case <id> --mhtml <archive> --output <new-directory> [--target <all|word|excel|powerpoint|onenote|rtf|markdown>] [--expected-sha256 <hash>] [--max-css-rules <count>] [--require-clean-source]");
            return 2;
        }

        string repoRoot = FindRepoRoot();
        using JsonDocument selection = JsonDocument.Parse(await File.ReadAllTextAsync(Path.Combine(repoRoot,
            "OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/html-h10-page-selection.json")));
        JsonElement page = selection.RootElement.GetProperty("pages").EnumerateArray().FirstOrDefault(candidate =>
            candidate.GetProperty("id").GetString() == caseId);
        if (page.ValueKind == JsonValueKind.Undefined || !page.TryGetProperty("editableTargets", out JsonElement selectedTargets)) {
            Console.Error.WriteLine($"Case '{caseId}' has no predeclared editable targets.");
            return 2;
        }

        string[] declared = selectedTargets.EnumerateArray().Select(target =>
            target.GetProperty("target").GetString()!.ToLowerInvariant()).ToArray();
        string requested = options.GetValueOrDefault("target", "all").ToLowerInvariant();
        if ((requested != "all" && !SupportedTargets.Contains(requested)) ||
            (requested != "all" && !declared.Contains(requested)) ||
            declared.Except(SupportedTargets).Any()) {
            Console.Error.WriteLine("Target is unsupported or absent from the case selection.");
            return 2;
        }
        string[] targets = requested == "all" ? declared : [requested];
        string[] markers = page.GetProperty("requiredVisibleMarkers").EnumerateArray()
            .Select(marker => marker.GetString()!).ToArray();
        if (markers.Length == 0 || markers.Any(string.IsNullOrWhiteSpace)) {
            Console.Error.WriteLine("The selected case needs nonempty predeclared visible markers.");
            return 2;
        }

        archivePath = Path.GetFullPath(archivePath);
        outputPath = Path.GetFullPath(outputPath);
        if (!File.Exists(archivePath) || Directory.Exists(outputPath) || File.Exists(outputPath)) {
            Console.Error.WriteLine("The source archive must exist and the output path must be new.");
            return 2;
        }
        long maximumArchiveBytes = selection.RootElement.GetProperty("capture").GetProperty("maximumArchiveBytes").GetInt64();
        if (new FileInfo(archivePath).Length > maximumArchiveBytes) {
            Console.Error.WriteLine($"The archive exceeds the selected corpus limit of {maximumArchiveBytes} bytes.");
            return 2;
        }
        byte[] archiveBytes = await File.ReadAllBytesAsync(archivePath);
        string sourceSha = Convert.ToHexString(SHA256.HashData(archiveBytes)).ToLowerInvariant();
        if (options.TryGetValue("expected-sha256", out string? expectedSha) &&
            !sourceSha.Equals(expectedSha, StringComparison.OrdinalIgnoreCase)) {
            Console.Error.WriteLine($"Archive SHA-256 mismatch: {sourceSha}");
            return 2;
        }
        int maxCssRules = 10000;
        if (options.TryGetValue("max-css-rules", out string? cssLimit) &&
            (!int.TryParse(cssLimit, out maxCssRules) || maxCssRules < 1 || maxCssRules > 40000)) {
            Console.Error.WriteLine("--max-css-rules must be between 1 and 40000.");
            return 2;
        }

        string commit = RunGit(repoRoot, "rev-parse", "HEAD").Trim();
        bool sourceClean = string.IsNullOrWhiteSpace(RunGit(repoRoot, "status", "--porcelain", "--untracked-files=normal"));
        if (options.ContainsKey("require-clean-source") && !sourceClean) {
            Console.Error.WriteLine("The source worktree is dirty; commit the runner before exact-head evidence.");
            return 2;
        }

        Directory.CreateDirectory(outputPath);
        List<TargetEvidence> results = [];
        foreach (string target in targets) {
            string targetDirectory = Path.Combine(outputPath, target);
            Directory.CreateDirectory(targetDirectory);
            Stopwatch timer = Stopwatch.StartNew();
            long allocationStart = GC.GetTotalAllocatedBytes(precise: true);
            TargetEvidence evidence;
            try {
                HtmlConversionDocumentOptions htmlOptions = HtmlConversionDocumentOptions.CreateUntrustedProfile();
                htmlOptions.Limits.MaxCssRules = maxCssRules;
                MhtmlDocument archive = MhtmlDocument.Load(archivePath, htmlOptions: htmlOptions);
                MhtmlImageEmbeddingResult prepared = archive.CreateEmbeddedImageDocumentResult();
                HtmlRenderOptions visualOptions = new() {
                    Mode = target == "word" ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous,
                    ViewportWidth = 816,
                    ViewportHeight = 900
                };
                archive.ConfigureRenderOptions(visualOptions);
                HtmlVisibleContentResult visible = await prepared.Value.CreateVisibleContentDocumentResultAsync(visualOptions);
                EditableExport export = EditableTargetExporter.SaveAndReopen(visible.Value, target, targetDirectory, caseId);
                string reopenedPath = Path.Combine(targetDirectory, "reopened.html");
                await File.WriteAllTextAsync(reopenedPath, export.ReopenedHtml);
                string decoded = System.Net.WebUtility.HtmlDecode(export.ReopenedHtml);
                Dictionary<string, bool> markerResults = markers.ToDictionary(marker => marker,
                    marker => decoded.Contains(marker, StringComparison.Ordinal));
                var diagnostics = prepared.Report.Diagnostics.Concat(visible.Report.Diagnostics)
                    .Concat(export.Report.Diagnostics).Select(diagnostic => new {
                        diagnostic.Component, diagnostic.Code, diagnostic.Source, diagnostic.Detail,
                        diagnostic.Message, severity = diagnostic.Severity.ToString(),
                        lossKind = diagnostic.LossKind.ToString(), diagnostic.Provenance
                    }).ToArray();
                bool reportSucceeded = prepared.Report.Succeeded && visible.Report.Succeeded && export.Report.Succeeded;
                bool reportHasLoss = prepared.Report.HasLoss || visible.Report.HasLoss || export.Report.HasLoss;
                evidence = new TargetEvidence(target, reportSucceeded && markerResults.Values.All(value => value),
                    reportSucceeded, reportHasLoss, markerResults, export.Artifact, new FileInfo(export.Artifact).Length,
                    reopenedPath, prepared.EmbeddedResourceCount, prepared.EmbeddedResourceBytes,
                    visible.AppliedStylesheetCount, visible.OmittedElementCount, export.MarkdownTableRows,
                    export.MarkdownTableColumns, export.PictureCount, export.LinkedPictureCount,
                    timer.Elapsed.TotalMilliseconds, GC.GetTotalAllocatedBytes(precise: true) - allocationStart,
                    diagnostics, null);
            } catch (Exception exception) {
                evidence = new TargetEvidence(target, false, false, true, null, null, null, null,
                    null, null, null, null, null, null, null, null, timer.Elapsed.TotalMilliseconds,
                    GC.GetTotalAllocatedBytes(precise: true) - allocationStart, null, exception.ToString());
            }
            results.Add(evidence);
            string targetReport = Path.Combine(targetDirectory, "report.json");
            await File.WriteAllTextAsync(targetReport, JsonSerializer.Serialize(evidence, JsonOptions));
            Console.WriteLine($"{target}: {(evidence.Passed ? "markers and reopen passed" : "failed")}, loss={evidence.ReportHasLoss}; {targetReport}");
        }
        var summary = new {
            caseId, role = page.GetProperty("role").GetString(), sourceUrl = page.GetProperty("url").GetString(),
            archivePath, sourceSha256 = sourceSha, sourceCommit = commit, worktreeDirty = !sourceClean,
            maxCssRules, requiredVisibleMarkers = markers, targetCount = results.Count,
            passedTargets = results.Count(result => result.Passed),
            targetsRequiringVisualReview = results.Where(result => result.ReportHasLoss).Select(result => result.Target).ToArray(),
            results = results.Select(result => new { result.Target, result.Passed, result.ReportHasLoss, result.Error }).ToArray()
        };
        string summaryPath = Path.Combine(outputPath, "summary.json");
        await File.WriteAllTextAsync(summaryPath, JsonSerializer.Serialize(summary, JsonOptions));
        Console.WriteLine(summaryPath);
        return results.All(result => result.Passed) ? 0 : 1;
    }

    private static bool TryParse(string[] args, out Dictionary<string, string> options) {
        options = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < args.Length; index++) {
            string argument = args[index];
            if (!argument.StartsWith("--", StringComparison.Ordinal)) return false;
            string key = argument[2..];
            if (key == "require-clean-source") {
                if (!options.TryAdd(key, "true")) return false;
            } else if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal) ||
                       !options.TryAdd(key, args[++index])) {
                return false;
            }
        }
        return options.Keys.All(key => key is "case" or "mhtml" or "output" or "target" or
            "expected-sha256" or "max-css-rules" or "require-clean-source");
    }

    private static string FindRepoRoot() {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) return directory.FullName;
            directory = directory.Parent;
        }
        throw new InvalidOperationException("OfficeIMO repository root was not found.");
    }

    private static string RunGit(string repoRoot, params string[] arguments) {
        ProcessStartInfo start = new("git") { WorkingDirectory = repoRoot, RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Could not run git.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0) throw new InvalidOperationException(error);
        return output;
    }
}

internal sealed record TargetEvidence(
    string Target, bool Passed, bool ReportSucceeded, bool ReportHasLoss,
    Dictionary<string, bool>? Markers, string? Artifact, long? ArtifactBytes, string? ReopenedHtmlArtifact,
    int? EmbeddedResourceCount, long? EmbeddedResourceBytes, int? AppliedStylesheetCount,
    int? OmittedElementCount, int? MarkdownTableRows, int? MarkdownTableColumns,
    int? PictureCount, int? LinkedPictureCount, double ElapsedMs, long AllocatedBytes,
    object? Diagnostics, string? Error);
