using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Qualification;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Benchmarks;

internal static partial class HtmlQualificationBaselineRunner {
    private const string ReportSchema = "officeimo.html.qualification-baseline";
    private const string ReportSchemaVersion = "1.0";
    private static readonly UTF8Encoding Utf8WithoutBom = new(false);
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    internal static int Run(string[] args) {
        string? claimedOutput = null;
        try {
            string? corpusPath = GetOption(args, "--corpus");
            string outputDirectory = GetOption(args, "--output") ?? Path.Combine(
                ".benchmark-artifacts", "html", "qualification", DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture));
            string fullOutput = Path.GetFullPath(outputDirectory);
            RequireEmptyOutputDirectory(fullOutput);
            Directory.CreateDirectory(fullOutput);
            claimedOutput = fullOutput;
            string sourceRoot = ResolveSourceRoot(GetOption(args, "--source-root"));
            string sourceCommit = ResolveCommit(sourceRoot);
            string[] untrackedSourcePaths = ResolveUntrackedSourcePaths(sourceRoot);

            HtmlQualificationCorpus corpus = HtmlQualificationCorpus.Load(corpusPath);
            var failures = new List<string>();
            HtmlQualificationDocumentEvidence document = BuildDocumentEvidence(corpus, failures);
            var profiles = new List<HtmlQualificationProfileEvidence>(corpus.Manifest.Profiles.Count);
            foreach (HtmlQualificationProfile profile in corpus.Manifest.Profiles) {
                profiles.Add(BuildProfileEvidence(corpus, profile, fullOutput, failures));
            }
            HtmlQualificationCancellationEvidence cancellation = ObserveCancellation(corpus);
            if (!cancellation.PreCanceledRenderStopped || cancellation.ResolverCalls != 0 || cancellation.ProducedRender) {
                failures.Add("A pre-canceled render did not stop before resource resolution and output.");
            }

            HtmlQualificationBaselineReport report = new(
                ReportSchema,
                ReportSchemaVersion,
                DateTimeOffset.UtcNow,
                sourceCommit,
                ResolveTrackedSourceTreeDirty(sourceRoot),
                untrackedSourcePaths.Length,
                ResolveUntrackedSourceRoots(untrackedSourcePaths),
                BuildCorpusEvidence(corpus),
                BuildEnvironmentEvidence(),
                BuildProviderEvidence(),
                document,
                profiles,
                cancellation,
                failures);
            string reportPath = Path.Combine(fullOutput, "baseline.json");
            File.WriteAllText(reportPath, JsonSerializer.Serialize(report, JsonOptions), Utf8WithoutBom);
            Console.WriteLine("Wrote " + reportPath);
            Console.WriteLine($"Corpus {corpus.Manifest.CorpusId}: {profiles.Count} profiles, {failures.Count} failures.");
            foreach (string failure in failures) Console.Error.WriteLine("QUALIFICATION FAILURE: " + failure);
            return failures.Count == 0 ? 0 : 1;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            if (claimedOutput != null) {
                try {
                    File.WriteAllText(Path.Combine(claimedOutput, "failure.txt"), exception.ToString(), Utf8WithoutBom);
                } catch {
                    // Preserve the original failure when the evidence directory is also unavailable.
                }
            }
            return 1;
        }
    }

    private static HtmlQualificationDocumentEvidence BuildDocumentEvidence(
        HtmlQualificationCorpus corpus,
        ICollection<string> failures) {
        var stages = new List<HtmlQualificationStageObservation>();
        HtmlConversionDocument source = Observe(stages, "parse", () => HtmlConversionDocument.Parse(
            corpus.EntryHtml,
            new HtmlConversionDocumentOptions {
                BaseUri = corpus.BaseUri,
                Profile = HtmlConversionProfile.HighFidelityPrint,
                Trust = HtmlInputTrust.Untrusted,
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            }));
        OfficeIMO.Html.Dom.HtmlDocument owned = Observe(stages, "owned-dom", () => source.Document);
        int nodeCount = owned.Descendants().Count() + 1;
        int elementCount = owned.Descendants().Count(node => node is OfficeIMO.Html.Dom.HtmlElement);

        IReadOnlyList<HtmlQualificationSelectorEvidence> selectors = Observe(stages, "owned-query", () =>
            corpus.Manifest.SelectorExpectations.Select(expectation => {
                int count = owned.QuerySelectorAll(expectation.Selector).Count;
                bool passed = count == expectation.Count;
                if (!passed) failures.Add($"Selector {expectation.Selector} matched {count}; expected {expectation.Count}.");
                return new HtmlQualificationSelectorEvidence(expectation.Selector, expectation.Count, count, passed);
            }).ToArray());

        HtmlLogicalDocument logical = Observe(stages, "logical-projection", () => source.LogicalDocument);
        HtmlSemanticDocument semantic = Observe(stages, "semantic-projection", () => source.SemanticDocument);
        HtmlComputedStyleSummary styleSummary = Observe(stages, "source-style-summary", () => source.StyleSummary);
        HtmlResourceManifest resources = Observe(stages, "resource-plan", () => source.ResourceManifest);
        string normalized = Observe(stages, "normalized-html", () => source.NormalizedHtml);
        if (source.Diagnostics.Count > 0) {
            failures.Add("The frozen qualification document produced conversion diagnostics: " +
                string.Join(", ", source.Diagnostics.Select(diagnostic => diagnostic.Code)) + ".");
        }
        int semanticBlockCount = 0;
        int semanticTableBlockCount = 0;
        foreach (HtmlSemanticSection section in semantic.Sections) {
            CountSemanticBlocks(section.Blocks, ref semanticBlockCount, ref semanticTableBlockCount);
        }

        HtmlQualificationInputFile entry = corpus.Manifest.Files.Single(file =>
            string.Equals(file.Path.Replace('\\', '/'), corpus.Manifest.EntryPath, StringComparison.Ordinal));
        return new HtmlQualificationDocumentEvidence(
            (int)entry.Length,
            entry.Sha256,
            nodeCount,
            elementCount,
            selectors,
            logical.GetCounts().ToDictionary(pair => pair.Key.ToString(), pair => pair.Value, StringComparer.Ordinal),
            logical.Capabilities,
            semantic.Sections.Count,
            semanticBlockCount,
            semantic.RootTables.Count,
            semanticTableBlockCount,
            semantic.Resources.Count,
            styleSummary.StyledElementCount,
            styleSummary.PropertyNames.Count,
            resources.AllowedCount,
            resources.BlockedCount,
            Hash(normalized),
            ToDiagnostics(source.Diagnostics),
            stages);
    }

    private static HtmlQualificationProfileEvidence BuildProfileEvidence(
        HtmlQualificationCorpus corpus,
        HtmlQualificationProfile profile,
        string outputDirectory,
        ICollection<string> failures) {
        var stages = new List<HtmlQualificationStageObservation>();
        HtmlRenderOptions options = CreateOptions(corpus, profile, out ResourceTracker renderResources);
        IReadOnlyList<HtmlQualificationStyleEvidence> styles = Observe(stages, "resolved-style-probes", () =>
            BuildStyleEvidence(corpus, profile, options, failures));
        HtmlRenderDocument rendered = ObserveAsync(stages, "resource-style-layout", () =>
            HtmlRenderEngine.RenderAsync(
                HtmlConversionDocument.Parse(corpus.EntryHtml, new HtmlConversionDocumentOptions {
                    BaseUri = corpus.BaseUri,
                    Profile = profile.Mode == "paged" ? HtmlConversionProfile.HighFidelityPrint : HtmlConversionProfile.Document,
                    Trust = HtmlInputTrust.Untrusted,
                    UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                    ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
                }),
                options));
        bool passed = true;
        if (rendered.Pages.Count != profile.ExpectedPageCount) {
            failures.Add($"Profile {profile.Id} rendered {rendered.Pages.Count} pages; expected {profile.ExpectedPageCount}.");
            passed = false;
        }
        foreach (string marker in corpus.Manifest.TextMarkers) {
            if (rendered.Text.Contains(marker, StringComparison.Ordinal)) continue;
            failures.Add($"Profile {profile.Id} lost text marker: {marker}.");
            passed = false;
        }
        if (styles.Any(style => !style.Passed)) passed = false;

        string profileDirectory = ResolveContainedOutputDirectory(outputDirectory, profile.Id);
        Directory.CreateDirectory(profileDirectory);
        var artifacts = new List<HtmlQualificationArtifactEvidence>();
        var pages = new List<HtmlQualificationPageEvidence>(rendered.Pages.Count);
        Observe(stages, "drawing-svg-raster", () => {
            foreach (HtmlRenderPage page in rendered.Pages) {
                OfficeDrawing drawing = page.CreateDrawing();
                string svg = OfficeDrawingSvgExporter.ToSvg(drawing, 1D);
                byte[] png = OfficeDrawingRasterRenderer.ToPng(drawing, 1D, OfficeColor.White);
                string stem = "page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture);
                string svgPath = Path.Combine(profileDirectory, stem + ".svg");
                string pngPath = Path.Combine(profileDirectory, stem + ".png");
                File.WriteAllText(svgPath, svg, Utf8WithoutBom);
                File.WriteAllBytes(pngPath, png);
                artifacts.Add(Artifact(outputDirectory, svgPath, "image/svg+xml", page.PageNumber));
                artifacts.Add(Artifact(outputDirectory, pngPath, "image/png", page.PageNumber));
                pages.Add(new HtmlQualificationPageEvidence(
                    page.PageNumber,
                    page.Width,
                    page.Height,
                    page.Scene.Count,
                    page.Visuals.Count,
                    Hash(svg),
                    Hash(png)));
            }
            return true;
        });

        var combinedResources = new ResourceTracker(corpus);
        combinedResources.AddRange(renderResources.Entries);
        IReadOnlyList<HtmlQualificationResourceEvidence> renderResourceEvidence = DistinctResources(renderResources.Entries);
        IReadOnlyList<HtmlQualificationPdfWarningEvidence> pdfWarnings = Array.Empty<HtmlQualificationPdfWarningEvidence>();
        IReadOnlyList<HtmlQualificationResourceEvidence> pdfResourceEvidence = Array.Empty<HtmlQualificationResourceEvidence>();
        HtmlQualificationPdfResourcePolicyEvidence? pdfResourcePolicyEvidence = null;
        string[] expectedResourcePaths = ExpectedResourcePaths(corpus);
        if (!ValidateResourceSet(profile.Id, "render", renderResourceEvidence, expectedResourcePaths, failures)) passed = false;
        if (profile.Mode == "paged") {
            PdfCore.PdfResourcePolicy resourcePolicy = CreateQualificationPdfResourcePolicy();
            pdfResourcePolicyEvidence = ToEvidence(resourcePolicy);
            HtmlToPdfOptions pdfOptions = new(CreateOptions(corpus, profile, out ResourceTracker pdfResources)) {
                ResourcePolicy = resourcePolicy
            };
            PdfCore.PdfDocumentConversionResult conversion = ObserveAsync(stages, "pdf-conversion", () =>
                HtmlConversionDocument.Parse(corpus.EntryHtml, new HtmlConversionDocumentOptions {
                    BaseUri = corpus.BaseUri,
                    Profile = HtmlConversionProfile.HighFidelityPrint,
                    Trust = HtmlInputTrust.Untrusted,
                    UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                    ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
                }).ToPdfDocumentResultAsync(pdfOptions));
            byte[] pdf = Observe(stages, "pdf-serialization", () => conversion.ToBytes());
            pdfWarnings = conversion.Warnings.Select(warning => new HtmlQualificationPdfWarningEvidence(
                warning.Converter,
                warning.Code,
                warning.Severity.ToString(),
                warning.LossKind.ToString(),
                warning.Message,
                warning.Source)).ToArray();
            if (conversion.HasLoss) {
                failures.Add($"Profile {profile.Id} PDF conversion reported loss: " +
                    string.Join(", ", conversion.Warnings.Select(warning => warning.Code)) + ".");
                passed = false;
            }
            combinedResources.AddRange(pdfResources.Entries);
            pdfResourceEvidence = DistinctResources(pdfResources.Entries);
            if (!ValidateResourceSet(profile.Id, "PDF", pdfResourceEvidence, expectedResourcePaths, failures)) passed = false;
            string pdfPath = Path.Combine(profileDirectory, "document.pdf");
            File.WriteAllBytes(pdfPath, pdf);
            artifacts.Add(Artifact(outputDirectory, pdfPath, "application/pdf", null));
            OfficeIMO.Pdf.PdfReadDocument readback = OfficeIMO.Pdf.PdfReadDocument.Open(pdf);
            string pdfText = readback.ExtractText();
            if (readback.Pages.Count != profile.ExpectedPageCount ||
                corpus.Manifest.TextMarkers.Any(marker => !ContainsNormalizedPhrase(pdfText, marker))) {
                failures.Add($"Profile {profile.Id} PDF readback did not preserve its page and searchable-text contract.");
                passed = false;
            }
        }

        if (rendered.HasLoss) {
            failures.Add($"Profile {profile.Id} reported rendering loss: " +
                string.Join(", ", rendered.Diagnostics.Select(diagnostic => diagnostic.Code)) + ".");
            passed = false;
        }

        IReadOnlyList<HtmlQualificationResourceEvidence> resourceEvidence = DistinctResources(combinedResources.Entries);
        return new HtmlQualificationProfileEvidence(
            profile.Id,
            profile.Mode,
            profile.Media,
            profile.ViewportWidth,
            profile.ViewportHeight,
            profile.ExpectedPageCount,
            rendered.Pages.Count,
            rendered.Text.Length,
            Hash(rendered.Text),
            rendered.Headings.Count,
            resourceEvidence.Count,
            resourceEvidence.Sum(resource => resource.Length),
            styles,
            resourceEvidence,
            renderResourceEvidence,
            pdfResourceEvidence,
            pdfResourcePolicyEvidence,
            pages,
            artifacts,
            ToDiagnostics(rendered.Diagnostics),
            pdfWarnings,
            stages,
            passed);
    }

    private static IReadOnlyList<HtmlQualificationStyleEvidence> BuildStyleEvidence(
        HtmlQualificationCorpus corpus,
        HtmlQualificationProfile profile,
        HtmlRenderOptions options,
        ICollection<string> failures) {
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(corpus.EntryHtml);
        string[] stylesheets = corpus.Manifest.Files
            .Where(file => file.Role == "stylesheet")
            .Select(file => Encoding.UTF8.GetString(corpus.ReadBytes(file)))
            .ToArray();
        HtmlRenderAdditionalStylesheetApplier.Apply(document, stylesheets);
        HtmlComputedStyleSet computed = HtmlComputedStyleEngine.ComputeForRendering(
            document,
            options,
            HtmlConversionLimits.CreateUntrustedProfile());
        var evidence = new List<HtmlQualificationStyleEvidence>(profile.StyleExpectations.Count);
        foreach (HtmlQualificationStyleExpectation expectation in profile.StyleExpectations) {
            IElement? element = document.QuerySelector(expectation.Selector);
            string value = string.Empty;
            HtmlQualificationCascadeEvidence? cascade = null;
            if (element != null && computed.Elements.TryGetValue(element, out HtmlComputedStyle? style)) {
                value = style.GetValue(expectation.Property);
                if (style.TryGetCascadePriority(expectation.Property, out HtmlCssCascadePriority priority)) {
                    cascade = new HtmlQualificationCascadeEvidence(
                        priority.Inherited,
                        priority.Important,
                        priority.Inline,
                        priority.LayerOrder != null,
                        priority.Ids,
                        priority.Classes,
                        priority.Elements,
                        priority.RuleOrder,
                        priority.DeclarationOrder);
                }
            }
            bool passed = value.Contains(expectation.Contains, StringComparison.OrdinalIgnoreCase);
            if (!passed) {
                failures.Add($"Profile {profile.Id} style {expectation.Selector} {expectation.Property} was '{value}'; expected it to contain '{expectation.Contains}'.");
            }
            evidence.Add(new HtmlQualificationStyleEvidence(
                expectation.Selector,
                expectation.Property,
                value,
                expectation.Contains,
                cascade,
                passed));
        }
        return evidence;
    }

    private static HtmlRenderOptions CreateOptions(
        HtmlQualificationCorpus corpus,
        HtmlQualificationProfile profile,
        out ResourceTracker tracker) {
        tracker = new ResourceTracker(corpus);
        ResourceTracker captured = tracker;
        var options = new HtmlRenderOptions {
            Mode = profile.Mode == "paged" ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous,
            ViewportWidth = profile.ViewportWidth,
            ViewportHeight = profile.ViewportHeight,
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(40D),
            BackgroundColor = OfficeColor.White,
            BaseUri = corpus.BaseUri,
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            MaxResourceCount = 16,
            MaxResourceBytes = 2L * 1024L * 1024L,
            MaxTotalResourceBytes = 4L * 1024L * 1024L,
            ResourceTimeout = TimeSpan.FromSeconds(10),
            ResourceResolver = captured.ResolveAsync
        };
        return options;
    }

    private static HtmlQualificationCancellationEvidence ObserveCancellation(HtmlQualificationCorpus corpus) {
        var tracker = new ResourceTracker(corpus);
        var options = new HtmlRenderOptions {
            BaseUri = corpus.BaseUri,
            ResourceResolver = tracker.ResolveAsync,
            ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        bool stopped = false;
        bool producedRender = false;
        try {
            HtmlRenderEngine.RenderAsync(HtmlConversionDocument.Parse(corpus.EntryHtml), options, cancellation.Token)
                .GetAwaiter().GetResult();
            producedRender = true;
        } catch (OperationCanceledException) {
            stopped = true;
        }
        return new HtmlQualificationCancellationEvidence(stopped, tracker.CallCount, producedRender);
    }

    private static HtmlQualificationCorpusEvidence BuildCorpusEvidence(HtmlQualificationCorpus corpus) => new(
        corpus.Manifest.CorpusId,
        corpus.ManifestSha256,
        corpus.Manifest.EntryPath,
        corpus.Manifest.BaseUri,
        corpus.Manifest.SourceKind,
        corpus.Manifest.License,
        corpus.Manifest.Files.Select(file => new HtmlQualificationFileEvidence(
            file.Path, file.Role, file.MediaType, file.Length, file.Sha256)).ToArray());

    private static HtmlQualificationArtifactEvidence Artifact(
        string outputRoot,
        string path,
        string mediaType,
        int? pageNumber) {
        byte[] bytes = File.ReadAllBytes(path);
        return new HtmlQualificationArtifactEvidence(
            GetRelativePath(outputRoot, path).Replace('\\', '/'),
            mediaType,
            bytes.LongLength,
            Hash(bytes),
            pageNumber);
    }

    private static IReadOnlyList<HtmlQualificationDiagnosticEvidence> ToDiagnostics(IEnumerable<HtmlDiagnostic> diagnostics) =>
        diagnostics.Select(diagnostic => new HtmlQualificationDiagnosticEvidence(
            diagnostic.Code,
            diagnostic.Severity.ToString(),
            diagnostic.LossKind.ToString(),
            diagnostic.Message,
            diagnostic.Source)).ToArray();

    private static T Observe<T>(ICollection<HtmlQualificationStageObservation> stages, string stage, Func<T> operation) {
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        T value = operation();
        stopwatch.Stop();
        stages.Add(new HtmlQualificationStageObservation(
            stage,
            "single-run-observation",
            stopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore));
        return value;
    }

    private static T ObserveAsync<T>(ICollection<HtmlQualificationStageObservation> stages, string stage, Func<Task<T>> operation) =>
        Observe(stages, stage, () => operation().GetAwaiter().GetResult());

    private static void CountSemanticBlocks(
        IEnumerable<HtmlSemanticBlock> blocks,
        ref int blockCount,
        ref int tableCount) {
        foreach (HtmlSemanticBlock block in blocks) {
            blockCount++;
            if (block.Table != null) tableCount++;
            CountSemanticBlocks(block.Children, ref blockCount, ref tableCount);
        }
    }

    private static bool ContainsNormalizedPhrase(string text, string marker) =>
        NormalizeWhitespace(text).Contains(NormalizeWhitespace(marker), StringComparison.Ordinal);

    private static string NormalizeWhitespace(string value) =>
        string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static void RequireEmptyOutputDirectory(string path) {
        if (Directory.Exists(path) && Directory.EnumerateFileSystemEntries(path).Any())
            throw new IOException("Qualification output directory must be empty: " + path + ".");
    }

    private static string ResolveContainedOutputDirectory(string outputRoot, string profileId) {
        string root = Path.GetFullPath(outputRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string candidate = Path.GetFullPath(Path.Combine(root, profileId));
        string prefix = root + Path.DirectorySeparatorChar;
        if (!candidate.StartsWith(prefix, PathComparison()))
            throw new InvalidDataException("Qualification profile output escaped its root: " + profileId + ".");
        return candidate;
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, argument => string.Equals(argument, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 >= args.Length) throw new ArgumentException(name + " requires a value.");
        return args[index + 1];
    }

    private static string[] ExpectedResourcePaths(HtmlQualificationCorpus corpus) => corpus.Manifest.Files
        .Where(file => file.Role is "stylesheet" or "font" or "image")
        .Select(file => file.Path)
        .OrderBy(path => path, StringComparer.Ordinal)
        .ToArray();

    private static IReadOnlyList<HtmlQualificationResourceEvidence> DistinctResources(
        IEnumerable<HtmlQualificationResourceEvidence> resources) => resources
        .GroupBy(resource => resource.Uri, StringComparer.Ordinal)
        .Select(group => group.First())
        .OrderBy(resource => resource.Uri, StringComparer.Ordinal)
        .ToArray();

    private static bool ValidateResourceSet(
        string profileId,
        string lane,
        IReadOnlyList<HtmlQualificationResourceEvidence> resources,
        string[] expectedPaths,
        ICollection<string> failures) {
        string[] actualPaths = resources.Select(resource => resource.Path)
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        if (actualPaths.SequenceEqual(expectedPaths, StringComparer.Ordinal)) return true;
        failures.Add($"Profile {profileId} {lane} lane resolved [{string.Join(", ", actualPaths)}]; expected [{string.Join(", ", expectedPaths)}].");
        return false;
    }

    private static PdfCore.PdfResourcePolicy CreateQualificationPdfResourcePolicy() => new() {
        AllowRemoteResourceResolution = true,
        AllowSystemFontEmbedding = false,
        AllowDocumentFontEmbedding = false,
        AllowLocalFileAccess = false,
        AllowDataUris = false,
        AllowEmbeddedPackageResources = false
    };

    private static HtmlQualificationPdfResourcePolicyEvidence ToEvidence(PdfCore.PdfResourcePolicy policy) => new(
        policy.AllowSystemFontEmbedding,
        policy.AllowDocumentFontEmbedding,
        policy.AllowLocalFileAccess,
        policy.AllowRemoteResourceResolution,
        policy.AllowDataUris,
        policy.AllowEmbeddedPackageResources);

    private static string GetRelativePath(string root, string path) {
        var rootUri = new Uri(root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar);
        return Uri.UnescapeDataString(rootUri.MakeRelativeUri(new Uri(path)).ToString()).Replace('/', Path.DirectorySeparatorChar);
    }

    private static string Hash(string value) => Hash(Encoding.UTF8.GetBytes(value));

    private static string Hash(byte[] bytes) {
        using SHA256 sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
    }

    private sealed class ResourceTracker {
        private readonly HtmlQualificationCorpus _corpus;
        private readonly ConcurrentQueue<HtmlQualificationResourceEvidence> _entries = new();
        private int _callCount;

        internal ResourceTracker(HtmlQualificationCorpus corpus) {
            _corpus = corpus;
        }

        internal int CallCount => Volatile.Read(ref _callCount);
        internal IReadOnlyList<HtmlQualificationResourceEvidence> Entries => _entries.ToArray();

        internal Task<HtmlResolvedResource?> ResolveAsync(HtmlRenderResourceRequest request, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            Interlocked.Increment(ref _callCount);
            if (!_corpus.TryResolve(request.Uri, out HtmlQualificationInputFile file, out byte[] bytes)) {
                throw new InvalidOperationException("Qualification resolver received an undeclared resource: " + request.Uri.AbsoluteUri + ".");
            }
            _entries.Enqueue(new HtmlQualificationResourceEvidence(
                request.Uri.AbsoluteUri,
                request.Kind.ToString(),
                file.Path,
                file.MediaType,
                file.Length,
                file.Sha256));
            return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(bytes, file.MediaType, request.Uri, 0));
        }

        internal void AddRange(IEnumerable<HtmlQualificationResourceEvidence> entries) {
            foreach (HtmlQualificationResourceEvidence entry in entries) _entries.Enqueue(entry);
        }
    }
}
