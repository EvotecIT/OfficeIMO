using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Css;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Benchmarks;

internal static class HtmlOwnedDocumentEvidenceRunner {
    private static readonly string[] Operations = ["Parse", "Query", "Edit", "Serialize", "ConversionOwned", "CssSyntax", "Cancel"];
    private static readonly (string Name, int Rows)[] Scales = [("Small", 10), ("Normal", 100), ("Large", 1000)];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 2) {
            Console.Error.WriteLine("Usage: --owned-document-evidence-probe <operation> <Small|Normal|Large>");
            return 2;
        }
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(ResolveOperation(args[0]), ResolveScale(args[1])), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int Run(string[] args, bool verifyBudgets) {
        try {
            string? operationFilter = GetOption(args, "--operation");
            string? scaleFilter = GetOption(args, "--scale");
            string? jsonPath = GetOption(args, "--json");
            int repeat = GetPositiveIntOption(args, "--repeat", 3);
            string[] operations = operationFilter == null ? Operations : [ResolveOperation(operationFilter)];
            (string Name, int Rows)[] scales = scaleFilter == null ? Scales : [ResolveScale(scaleFilter)];
            HtmlOwnedDocumentBudgetManifest? manifest = verifyBudgets ? LoadBudgetManifest() : null;
            var measurements = new List<HtmlOwnedDocumentEvidenceMeasurement>();
            var failures = new List<string>();
            foreach (var scale in scales) {
                foreach (string operation in operations) {
                    for (int iteration = 1; iteration <= repeat; iteration++) {
                        HtmlOwnedDocumentEvidenceMeasurement measurement = RunChildProbe(operation, scale.Name) with { Iteration = iteration };
                        measurements.Add(measurement);
                        Console.WriteLine(
                            $"{operation,-16} {scale.Name,-6} #{iteration,-2} " +
                            $"{measurement.ElapsedMilliseconds,9:F3} ms " +
                            $"{measurement.AllocatedBytes / 1024D,10:F1} KiB alloc " +
                            $"{measurement.RetainedManagedHeapGrowthBytes / 1024D,10:F1} KiB retained " +
                            $"{measurement.PeakManagedHeapGrowthBytes / 1024D,10:F1} KiB managed peak");
                    }
                }
            }
            if (manifest != null) EvaluateBudgets(manifest, measurements, failures);
            var report = new HtmlOwnedDocumentEvidenceReport(
                DateTimeOffset.UtcNow, ResolveCommit(), ResolveTrackedSourceDirty(),
                RuntimeInformation.FrameworkDescription, RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(), Environment.ProcessorCount,
                repeat, measurements, failures);
            if (jsonPath != null) {
                string fullPath = Path.GetFullPath(jsonPath);
                string? directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
                File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
                Console.WriteLine("Wrote " + fullPath);
            }
            foreach (string failure in failures) Console.Error.WriteLine("BUDGET FAILURE: " + failure);
            return failures.Count == 0 ? 0 : 1;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static HtmlOwnedDocumentEvidenceMeasurement Measure(string operation, (string Name, int Rows) scale) {
        EvidenceOperation evidence = EvidenceOperation.Create(operation, scale);
        for (int index = 0; index < 2; index++) evidence.Validate(evidence.Execute());
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        using var sampler = new ManagedHeapSampler();
        var stopwatch = Stopwatch.StartNew();
        object result = evidence.Execute();
        stopwatch.Stop();
        long peakManaged = sampler.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        EvidenceValidation validation = evidence.Validate(result);
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        GC.KeepAlive(result);
        GC.KeepAlive(evidence);
        return new HtmlOwnedDocumentEvidenceMeasurement(
            operation, scale.Name, 1, evidence.InputCharacters, validation.ResultItems,
            validation.FingerprintSha256, stopwatch.Elapsed.TotalMilliseconds, allocated, retained,
            Math.Max(0, peakManaged - heapBefore), process.PeakWorkingSet64, validation.OutputCharacters);
    }

    private static HtmlOwnedDocumentEvidenceMeasurement RunChildProbe(string operation, string scale) {
        string processPath = Environment.ProcessPath ?? throw new InvalidOperationException("Unable to resolve benchmark process path.");
        var info = new ProcessStartInfo {
            FileName = processPath,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase))
            info.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        info.ArgumentList.Add("--owned-document-evidence-probe");
        info.ArgumentList.Add(operation);
        info.ArgumentList.Add(scale);
        using Process child = Process.Start(info) ?? throw new InvalidOperationException("Unable to start owned-document probe.");
        Task<string> outputTask = child.StandardOutput.ReadToEndAsync();
        Task<string> errorTask = child.StandardError.ReadToEndAsync();
        child.WaitForExit();
        string output = outputTask.GetAwaiter().GetResult();
        string error = errorTask.GetAwaiter().GetResult();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe {operation}/{scale} failed: {error}");
        return JsonSerializer.Deserialize<HtmlOwnedDocumentEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {operation}/{scale} returned no measurement.");
    }

    private static void EvaluateBudgets(
        HtmlOwnedDocumentBudgetManifest manifest,
        IReadOnlyList<HtmlOwnedDocumentEvidenceMeasurement> measurements,
        ICollection<string> failures) {
        foreach (IGrouping<(string Operation, string Scale), HtmlOwnedDocumentEvidenceMeasurement> group in measurements.GroupBy(
                     value => (value.Operation, value.Scale))) {
            HtmlOwnedDocumentBudget? budget = manifest.Budgets.FirstOrDefault(value =>
                string.Equals(value.Operation, group.Key.Operation, StringComparison.OrdinalIgnoreCase)
                && string.Equals(value.Scale, group.Key.Scale, StringComparison.OrdinalIgnoreCase));
            if (budget == null) {
                failures.Add($"Missing budget for {group.Key.Operation}/{group.Key.Scale}.");
                continue;
            }
            string lane = group.Key.Operation + "/" + group.Key.Scale;
            Check(Median(group.Select(value => value.ElapsedMilliseconds)), budget.MaxElapsedMilliseconds, "median elapsed ms");
            Check(Median(group.Select(value => (double)value.AllocatedBytes)), budget.MaxAllocatedBytes, "median allocated bytes");
            Check(Median(group.Select(value => (double)value.RetainedManagedHeapGrowthBytes)), budget.MaxRetainedManagedHeapGrowthBytes, "median retained bytes");
            Check(group.Max(value => (double)value.PeakManagedHeapGrowthBytes), budget.MaxPeakManagedHeapGrowthBytes, "maximum managed peak bytes");
            Check(group.Max(value => (double)value.AbsoluteProcessPeakWorkingSetBytes), budget.MaxAbsoluteProcessPeakWorkingSetBytes, "maximum process peak bytes");
            Check(group.Max(value => (double)value.OutputCharacters), budget.MaxOutputCharacters, "maximum output characters");
            void Check(double actual, double maximum, string metric) {
                if (actual > maximum) failures.Add($"{lane}: {metric} {actual:F0} > {maximum:F0}.");
            }
        }
    }

    private static HtmlOwnedDocumentBudgetManifest LoadBudgetManifest() {
        string path = Path.Combine(AppContext.BaseDirectory, "html-owned-document-performance-budgets.json");
        return JsonSerializer.Deserialize<HtmlOwnedDocumentBudgetManifest>(File.ReadAllText(path), JsonOptions)
            ?? throw new InvalidOperationException("Owned-document performance budget manifest is invalid.");
    }

    private static string ResolveOperation(string value) => Operations.FirstOrDefault(
        operation => string.Equals(operation, value, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException("Unknown owned-document operation: " + value);
    private static (string Name, int Rows) ResolveScale(string value) => Scales.FirstOrDefault(
        scale => string.Equals(scale.Name, value, StringComparison.OrdinalIgnoreCase)) is var resolved && resolved.Rows > 0
            ? resolved : throw new ArgumentException("Unknown owned-document scale: " + value);
    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal)) throw new ArgumentException(name + " requires a value.");
        return args[index + 1];
    }
    private static int GetPositiveIntOption(string[] args, string name, int defaultValue) {
        string? value = GetOption(args, name);
        return value == null ? defaultValue : int.TryParse(value, out int parsed) && parsed > 0
            ? parsed : throw new ArgumentException(name + " must be positive.");
    }
    private static double Median(IEnumerable<double> values) {
        double[] ordered = values.OrderBy(value => value).ToArray();
        int middle = ordered.Length / 2;
        return ordered.Length % 2 == 0 ? (ordered[middle - 1] + ordered[middle]) / 2D : ordered[middle];
    }
    private static string ResolveCommit() {
        string? environment = Environment.GetEnvironmentVariable("GITHUB_SHA");
        return string.IsNullOrWhiteSpace(environment) ? RunGit("rev-parse", "HEAD") ?? "unknown" : environment;
    }
    private static bool ResolveTrackedSourceDirty() => RunGit("diff", "--quiet", "HEAD", "--") == null;
    private static string? RunGit(params string[] arguments) {
        try {
            var info = new ProcessStartInfo("git") {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            foreach (string argument in arguments) info.ArgumentList.Add(argument);
            using Process process = Process.Start(info)!;
            Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
            Task<string> errorTask = process.StandardError.ReadToEndAsync();
            process.WaitForExit();
            string output = outputTask.GetAwaiter().GetResult().Trim();
            _ = errorTask.GetAwaiter().GetResult();
            return process.ExitCode == 0 ? output : null;
        } catch { return null; }
    }

    private sealed class EvidenceOperation {
        private readonly Func<object> _execute;
        private readonly Func<object, EvidenceValidation> _validate;
        private EvidenceOperation(int inputCharacters, Func<object> execute, Func<object, EvidenceValidation> validate) {
            InputCharacters = inputCharacters; _execute = execute; _validate = validate;
        }
        internal int InputCharacters { get; }
        internal object Execute() => _execute();
        internal EvidenceValidation Validate(object result) => _validate(result);

        internal static EvidenceOperation Create(string operation, (string Name, int Rows) scale) {
            if (operation == "Cancel") return CreateInFlightCancellation(scale.Name);
            int rows = scale.Rows;
            string html = HtmlBenchmarkCorpus.BuildReport(rows);
            int elementCount = checked(rows * 4 + 20);
            return operation switch {
                "Parse" => new EvidenceOperation(html.Length,
                    () => HtmlDocumentEngine.Default.ParseDocument(html),
                    result => ValidateDocument((HtmlDocument)result, elementCount)),
                "Query" => CreateQuery(html, rows),
                "Edit" => CreateEdit(html, elementCount),
                "Serialize" => CreateSerialize(html, elementCount),
                "ConversionOwned" => new EvidenceOperation(html.Length,
                    () => ParseConversionOwned(html),
                    result => ValidateConversion((HtmlConversionDocument)result, elementCount)),
                "CssSyntax" => CreateCssSyntax(rows),
                _ => throw new ArgumentOutOfRangeException(nameof(operation))
            };
        }

        private static EvidenceOperation CreateInFlightCancellation(string scale) {
            int rows = scale switch {
                "Small" => 10_000,
                "Normal" => 25_000,
                "Large" => 100_000,
                _ => throw new ArgumentOutOfRangeException(nameof(scale))
            };
            string html = HtmlBenchmarkCorpus.BuildReport(rows);
            return new EvidenceOperation(html.Length,
                () => ExecuteInFlightCanceledParse(html),
                result => new EvidenceValidation((int)result, Hash("cancelled-in-flight"), 0));
        }

        private static EvidenceOperation CreateQuery(string html, int rows) {
            HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(html);
            int expected = (rows + 1) / 2 + 2;
            return new EvidenceOperation(html.Length,
                () => document.QuerySelectorAll("tbody tr:nth-child(odd), section.card"),
                result => {
                    var matches = (IReadOnlyList<HtmlElement>)result;
                    if (matches.Count != expected) throw new InvalidOperationException($"Query returned {matches.Count}; expected {expected}.");
                    return new EvidenceValidation(matches.Count, Hash(string.Join("|", matches.Select(value => value.LocalName + ":" + value.TextContent))), 0);
                });
        }

        private static EvidenceOperation CreateEdit(string html, int elementCount) {
            HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(html);
            return new EvidenceOperation(html.Length,
                () => document.Edit(editor => {
                    HtmlElement heading = editor.Children.SelectMany(DescendantElements).First(value => value.LocalName == "h1");
                    heading.TextContent = "Edited benchmark report";
                    heading.SetAttribute("data-evidence", "owned");
                }),
                result => {
                    var edited = (HtmlDocument)result;
                    HtmlElement heading = edited.Descendants().OfType<HtmlElement>().First(value => value.LocalName == "h1");
                    if (heading.TextContent != "Edited benchmark report" || heading.GetAttribute("data-evidence") != "owned")
                        throw new InvalidOperationException("Owned edit did not preserve the requested mutation.");
                    return ValidateDocument(edited, elementCount);
                });
        }

        private static EvidenceOperation CreateSerialize(string html, int elementCount) {
            HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(html);
            return new EvidenceOperation(html.Length,
                () => document.OuterHtml,
                result => {
                    string output = (string)result;
                    HtmlDocument reopened = HtmlDocumentEngine.Default.ParseDocument(output);
                    EvidenceValidation validation = ValidateDocument(reopened, elementCount);
                    return validation with { OutputCharacters = output.Length, FingerprintSha256 = Hash(output) };
                });
        }

        private static EvidenceOperation CreateCssSyntax(int rows) {
            (string css, _) = HtmlProviderBenchmarkCorpus.BuildStyledCards(rows);
            return new EvidenceOperation(css.Length,
                () => HtmlCssSyntaxParser.ParseStyleSheet(css),
                result => {
                    var sheet = (HtmlCssStyleSheet)result;
                    if (!string.Equals(sheet.Source, css, StringComparison.Ordinal) || sheet.Rules.Count != rows + 5)
                        throw new InvalidOperationException($"Owned CSS syntax retained {sheet.Rules.Count} top-level rules; expected {rows + 5}.");
                    return new EvidenceValidation(sheet.Rules.Count, Hash(sheet.ToCss()), sheet.Source.Length);
                });
        }

        private static HtmlConversionDocument ParseConversionOwned(string html) {
            HtmlConversionDocument conversion = HtmlConversionDocument.Parse(html);
            _ = conversion.Document;
            return conversion;
        }

        private static int ExecuteInFlightCanceledParse(string html) {
            using var cancellation = new CancellationTokenSource();
            using var cancellationThreadReady = new ManualResetEventSlim();
            using var providerEntered = new ManualResetEventSlim();
            var engine = new HtmlDocumentEngine(new SignalingParserProvider(
                HtmlDocumentEngine.Default.ParserProvider, providerEntered));
            var cancellationThread = new Thread(() => {
                cancellationThreadReady.Set();
                providerEntered.Wait();
                Thread.Sleep(5);
                cancellation.Cancel();
            }) { IsBackground = true, Name = "OfficeIMO.Html in-flight cancellation evidence" };
            cancellationThread.Start();
            cancellationThreadReady.Wait();
            var stopwatch = Stopwatch.StartNew();
            try {
                engine.ParseDocument(html, new HtmlParseOptions {
                    MaxInputCharacters = html.Length,
                    MaxNodes = null
                }, cancellation.Token);
            } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
                stopwatch.Stop();
                cancellationThread.Join();
                if (stopwatch.ElapsedMilliseconds < 2)
                    throw new InvalidOperationException("Cancellation occurred before the parse performed measurable work.");
                return 1;
            }
            cancellationThread.Join();
            throw new InvalidOperationException("The parse completed before the in-flight cancellation was observed.");
        }

        private sealed class SignalingParserProvider : IHtmlParserProvider {
            private readonly IHtmlParserProvider _inner;
            private readonly ManualResetEventSlim _entered;

            internal SignalingParserProvider(IHtmlParserProvider inner, ManualResetEventSlim entered) {
                _inner = inner;
                _entered = entered;
            }

            public string Id => _inner.Id;

            public HtmlDocument ParseDocument(string source, HtmlParseOptions options, CancellationToken cancellationToken = default) {
                if (cancellationToken.IsCancellationRequested)
                    throw new InvalidOperationException("Cancellation was requested before the parser provider was entered.");
                _entered.Set();
                return _inner.ParseDocument(source, options, cancellationToken);
            }

            public HtmlDocumentFragment ParseFragment(
                string source,
                HtmlElement contextElement,
                HtmlParseOptions options,
                CancellationToken cancellationToken = default) =>
                _inner.ParseFragment(source, contextElement, options, cancellationToken);
        }

        private static EvidenceValidation ValidateConversion(HtmlConversionDocument conversion, int expectedElements) =>
            ValidateDocument(conversion.Document, expectedElements);
        private static EvidenceValidation ValidateDocument(HtmlDocument document, int expectedElements) {
            HtmlElement[] elements = document.Descendants().OfType<HtmlElement>().ToArray();
            if (elements.Length != expectedElements || elements.First(value => value.LocalName == "h1").TextContent is not ("Benchmark Report" or "Edited benchmark report"))
                throw new InvalidOperationException("Owned document evidence lost report structure.");
            string structural = string.Join("|", elements.Select(value => value.LocalName + ":" + value.Attributes.Count));
            return new EvidenceValidation(elements.Length, Hash(structural), 0);
        }
        private static IEnumerable<HtmlElement> DescendantElements(HtmlElement element) {
            yield return element;
            foreach (HtmlElement descendant in element.Descendants().OfType<HtmlElement>()) yield return descendant;
        }
    }

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    private readonly record struct EvidenceValidation(int ResultItems, string FingerprintSha256, int OutputCharacters);

    private sealed class ManagedHeapSampler : IDisposable {
        private readonly Thread _thread;
        private readonly ManualResetEventSlim _stop = new(false);
        private long _peak = GC.GetTotalMemory(forceFullCollection: false);
        private int _stopped;
        internal ManagedHeapSampler() {
            _thread = new Thread(Sample) { IsBackground = true, Name = "OfficeIMO.Html owned-document heap sampler" };
            _thread.Start();
        }
        internal long Stop() {
            if (Interlocked.Exchange(ref _stopped, 1) == 0) { _stop.Set(); _thread.Join(); Record(); }
            return Interlocked.Read(ref _peak);
        }
        public void Dispose() { Stop(); _stop.Dispose(); }
        private void Sample() { while (!_stop.Wait(1)) Record(); }
        private void Record() {
            long observed = GC.GetTotalMemory(forceFullCollection: false);
            long current = Interlocked.Read(ref _peak);
            while (observed > current) {
                long prior = Interlocked.CompareExchange(ref _peak, observed, current);
                if (prior == current) return;
                current = prior;
            }
        }
    }
}
