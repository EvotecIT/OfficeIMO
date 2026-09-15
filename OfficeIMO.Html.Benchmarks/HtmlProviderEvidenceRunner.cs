using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using AngleSharp.Css.Parser;
using OfficeIMO.Html.Dom;
using NativeDocumentType = AngleSharp.Dom.IDocumentType;
using NativeElement = AngleSharp.Dom.IElement;
using NativeNode = AngleSharp.Dom.INode;
using NativeHtmlDocument = AngleSharp.Html.Dom.IHtmlDocument;
using NativeTemplateElement = AngleSharp.Html.Dom.IHtmlTemplateElement;

namespace OfficeIMO.Html.Benchmarks;

internal static class HtmlProviderEvidenceRunner {
    private const int RetainedBatchSize = 8;
    private static readonly string[] Scenarios = [
        "AngleSharpHtmlNative",
        "OfficeIMODocumentOwned",
        "ConversionNative",
        "ConversionNativeAndOwned",
        "AngleSharpCssSyntax",
        "OfficeIMOCssSyntax",
        "OfficeIMOCssCascade"
    ];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 1) {
            Console.Error.WriteLine("Usage: --provider-evidence-probe <scenario>");
            return 2;
        }
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(args[0]), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int Run(string[] args) {
        try {
            string? scenarioFilter = GetOption(args, "--scenario");
            string? jsonPath = GetOption(args, "--json");
            int repeat = GetPositiveIntOption(args, "--repeat", 3);
            string[] scenarios = string.IsNullOrWhiteSpace(scenarioFilter)
                ? Scenarios
                : [ResolveScenario(scenarioFilter!)];
            var measurements = new List<HtmlProviderEvidenceMeasurement>(scenarios.Length * repeat);
            foreach (string scenario in scenarios) {
                for (int iteration = 1; iteration <= repeat; iteration++) {
                    HtmlProviderEvidenceMeasurement measurement = RunChildProbe(scenario) with { Iteration = iteration };
                    measurements.Add(measurement);
                    Console.WriteLine(
                        $"{scenario,-26} #{iteration} {measurement.ElapsedMilliseconds,9:F3} ms " +
                        $"{measurement.AllocatedBytes / 1024D,10:F1} KiB alloc " +
                        $"{measurement.RetainedManagedHeapBytesPerResult / 1024D,10:F1} KiB retained/result " +
                        $"{measurement.ResultItems,6:N0} items");
                }
            }
            ValidateEquivalentHtmlLanes(measurements);

            HtmlProviderEvidenceSummary[] summaries = scenarios.Select(scenario => {
                HtmlProviderEvidenceMeasurement[] values = measurements.Where(value => value.Scenario == scenario).ToArray();
                return new HtmlProviderEvidenceSummary(
                    scenario,
                    Median(values.Select(value => value.ElapsedMilliseconds)),
                    Median(values.Select(value => (double)value.AllocatedBytes)),
                    Median(values.Select(value => (double)value.RetainedManagedHeapBytesPerResult)),
                    Median(values.Select(value => (double)value.AbsoluteProcessPeakWorkingSetBytes)),
                    values[0].InputCharacters,
                    values[0].ResultItems);
            }).ToArray();
            var report = new HtmlProviderEvidenceReport(
                DateTimeOffset.UtcNow, ResolveCommit(), ResolveTrackedSourceDirty(),
                RuntimeInformation.FrameworkDescription, RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(), Environment.ProcessorCount,
                repeat, BuildProviderEvidence(), measurements, summaries);
            if (!string.IsNullOrWhiteSpace(jsonPath)) {
                string fullPath = Path.GetFullPath(jsonPath!);
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
                File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
                Console.WriteLine("Wrote " + fullPath);
            }
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static HtmlProviderEvidenceMeasurement Measure(string scenarioName) {
        string scenario = ResolveScenario(scenarioName);
        ProviderEvidenceOperation operation = ProviderEvidenceOperation.Create(scenario);
        for (int index = 0; index < 3; index++) {
            object warm = operation.Execute();
            _ = operation.Validate(warm);
            GC.KeepAlive(warm);
        }

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        object timed = operation.Execute();
        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        ProviderEvidenceValidation validation = operation.Validate(timed);
        GC.KeepAlive(timed);

        timed = null!;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        object[] retained = new object[RetainedBatchSize];
        for (int index = 0; index < retained.Length; index++) {
            retained[index] = operation.Execute();
            _ = operation.Validate(retained[index]);
        }
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retainedPerResult = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore) / retained.Length;
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long processPeak = process.PeakWorkingSet64;
        GC.KeepAlive(retained);

        return new HtmlProviderEvidenceMeasurement(
            scenario, 1, operation.InputCharacters, validation.ResultItems, validation.FingerprintSha256,
            stopwatch.Elapsed.TotalMilliseconds, allocated, retainedPerResult, processPeak);
    }

    private static HtmlProviderEvidenceMeasurement RunChildProbe(string scenario) {
        string processPath = Environment.ProcessPath ?? throw new InvalidOperationException("Unable to resolve process path.");
        var info = new ProcessStartInfo {
            FileName = processPath,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase)) {
            info.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        }
        info.ArgumentList.Add("--provider-evidence-probe");
        info.ArgumentList.Add(scenario);
        using Process child = Process.Start(info) ?? throw new InvalidOperationException("Unable to start HTML provider probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe {scenario} failed: {error}");
        return JsonSerializer.Deserialize<HtmlProviderEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {scenario} returned no measurement.");
    }

    private static string ResolveScenario(string value) => Scenarios.FirstOrDefault(
        scenario => string.Equals(scenario, value, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException("Unknown HTML provider scenario: " + value);

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, argument => string.Equals(argument, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 >= args.Length) throw new ArgumentException(name + " requires a value.");
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
        string? value = Environment.GetEnvironmentVariable("GITHUB_SHA");
        if (!string.IsNullOrWhiteSpace(value)) return value;
        try {
            using Process process = Process.Start(GitInfo("rev-parse", "HEAD"))!;
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 ? output : "unknown";
        } catch { return "unknown"; }
    }

    private static IReadOnlyList<HtmlProviderAssemblyEvidence> BuildProviderEvidence() =>
        new[] {
            typeof(AngleSharp.Html.Parser.HtmlParser).Assembly,
            typeof(CssParser).Assembly,
            typeof(OfficeIMO.Html.Providers.AngleSharpHtmlParser).Assembly,
            typeof(System.Text.CodePagesEncodingProvider).Assembly,
            typeof(HtmlDocumentEngine).Assembly,
            typeof(HtmlDocument).Assembly
        }
        .Distinct()
        .Select(assembly => new HtmlProviderAssemblyEvidence(
            assembly.GetName().Name ?? "unknown",
            assembly.GetName().Version?.ToString() ?? "unknown",
            assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "unknown",
            string.IsNullOrWhiteSpace(assembly.Location) || !File.Exists(assembly.Location)
                ? 0
                : new FileInfo(assembly.Location).Length))
        .OrderBy(item => item.Assembly, StringComparer.Ordinal)
        .ToArray();

    private static bool ResolveTrackedSourceDirty() {
        try {
            using Process tracked = Process.Start(GitInfo("diff", "--quiet", "HEAD", "--"))!;
            tracked.WaitForExit();
            return tracked.ExitCode != 0;
        } catch { return true; }
    }

    private static ProcessStartInfo GitInfo(params string[] arguments) {
        var info = new ProcessStartInfo("git") { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
        foreach (string argument in arguments) info.ArgumentList.Add(argument);
        return info;
    }

    private static void ValidateEquivalentHtmlLanes(IReadOnlyList<HtmlProviderEvidenceMeasurement> measurements) {
        string[] htmlScenarios = Scenarios.Take(4).ToArray();
        HtmlProviderEvidenceMeasurement[] html = measurements
            .Where(measurement => htmlScenarios.Contains(measurement.Scenario, StringComparer.Ordinal))
            .ToArray();
        if (html.Length == 0) return;
        HtmlProviderEvidenceMeasurement expected = html[0];
        foreach (HtmlProviderEvidenceMeasurement measurement in html) {
            if (measurement.ResultItems != expected.ResultItems
                || !string.Equals(measurement.ResultFingerprintSha256, expected.ResultFingerprintSha256, StringComparison.Ordinal)) {
                throw new InvalidOperationException(
                    $"HTML provider evidence is not structurally equivalent: {measurement.Scenario} iteration {measurement.Iteration} " +
                    $"returned {measurement.ResultItems} items and fingerprint {measurement.ResultFingerprintSha256}; " +
                    $"expected {expected.ResultItems} and {expected.ResultFingerprintSha256}.");
            }
        }
    }

    private sealed class ProviderEvidenceOperation {
        private readonly Func<object> _execute;
        private readonly Func<object, ProviderEvidenceValidation> _validate;

        private ProviderEvidenceOperation(int inputCharacters, Func<object> execute, Func<object, ProviderEvidenceValidation> validate) {
            InputCharacters = inputCharacters;
            _execute = execute;
            _validate = validate;
        }

        internal int InputCharacters { get; }
        internal object Execute() => _execute();
        internal ProviderEvidenceValidation Validate(object result) => _validate(result);

        internal static ProviderEvidenceOperation Create(string scenario) {
            string html = HtmlBenchmarkCorpus.BuildReport(100);
            (string css, string styledHtml) = HtmlProviderBenchmarkCorpus.BuildStyledCards(100);
            return scenario switch {
                "AngleSharpHtmlNative" => new ProviderEvidenceOperation(
                    html.Length,
                    () => HtmlDocumentParser.ParseDocument(html),
                    result => ValidateNative((NativeHtmlDocument)result)),
                "OfficeIMODocumentOwned" => new ProviderEvidenceOperation(
                    html.Length,
                    () => HtmlDocumentEngine.Default.ParseDocument(html),
                    result => ValidateOwned((HtmlDocument)result)),
                "ConversionNative" => new ProviderEvidenceOperation(
                    html.Length,
                    () => HtmlConversionDocument.Parse(html),
                    result => ValidateConversion((HtmlConversionDocument)result, requireOwned: false)),
                "ConversionNativeAndOwned" => new ProviderEvidenceOperation(
                    html.Length,
                    () => ParseConversionWithOwned(html),
                    result => ValidateConversion((HtmlConversionDocument)result, requireOwned: true)),
                "AngleSharpCssSyntax" => CreateCssSyntax(css),
                "OfficeIMOCssSyntax" => CreateOwnedCssSyntax(css),
                "OfficeIMOCssCascade" => CreateCssCascade(styledHtml),
                _ => throw new ArgumentOutOfRangeException(nameof(scenario))
            };
        }

        private static ProviderEvidenceOperation CreateCssSyntax(string css) {
            var parser = new CssParser(new CssParserOptions { IsIncludingUnknownDeclarations = true });
            return new ProviderEvidenceOperation(
                css.Length,
                () => parser.ParseStyleSheet(css),
                result => {
                    int count = ((AngleSharp.Css.Dom.ICssStyleSheet)result).Rules.Length;
                    if (count < 3) throw new InvalidOperationException("The CSS provider evidence lost top-level rules.");
                    return new ProviderEvidenceValidation(count, "not-applicable");
                });
        }

        private static ProviderEvidenceOperation CreateCssCascade(string html) {
            HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(html);
            return new ProviderEvidenceOperation(
                html.Length,
                () => HtmlComputedStyleEngine.Compute(document, HtmlCssMediaContext.Screen),
                result => {
                    int count = ((IReadOnlyDictionary<HtmlElement, HtmlComputedStyle>)result).Count;
                    if (count < 200) throw new InvalidOperationException("The owned CSS cascade lost styled-card elements.");
                    return new ProviderEvidenceValidation(count, "not-applicable");
                });
        }

        private static ProviderEvidenceOperation CreateOwnedCssSyntax(string css) {
            return new ProviderEvidenceOperation(
                css.Length,
                () => OfficeIMO.Html.Css.HtmlCssSyntaxParser.ParseStyleSheet(css),
                result => {
                    var sheet = (OfficeIMO.Html.Css.HtmlCssStyleSheet)result;
                    if (!string.Equals(sheet.ToCss(), css, StringComparison.Ordinal) || sheet.Rules.Count < 3)
                        throw new InvalidOperationException("The owned CSS syntax evidence lost source or top-level rules.");
                    return new ProviderEvidenceValidation(sheet.Rules.Count, Hash(new StringBuilder(sheet.ToCss())));
                });
        }

        private static HtmlConversionDocument ParseConversionWithOwned(string html) {
            HtmlConversionDocument conversion = HtmlConversionDocument.Parse(html);
            _ = conversion.Document;
            return conversion;
        }

        private static ProviderEvidenceValidation ValidateNative(NativeHtmlDocument document) {
            int count = document.QuerySelectorAll("*").Length;
            if (count != 420 || document.QuerySelector("h1")?.TextContent != "Benchmark Report")
                throw new InvalidOperationException("The native provider evidence lost report structure.");
            return new ProviderEvidenceValidation(count, FingerprintNative(document));
        }

        private static ProviderEvidenceValidation ValidateOwned(HtmlDocument document) {
            int count = document.QuerySelectorAll("*").Count;
            if (count != 420 || document.QuerySelector("h1")?.TextContent != "Benchmark Report")
                throw new InvalidOperationException("The owned provider evidence lost report structure.");
            return new ProviderEvidenceValidation(count, FingerprintOwned(document));
        }

        private static ProviderEvidenceValidation ValidateConversion(HtmlConversionDocument document, bool requireOwned) {
            if (document.SourceHtml.Length == 0) throw new InvalidOperationException("The conversion provider evidence lost its source.");
            return requireOwned
                ? ValidateOwned(document.Document)
                : document.ProjectSourceDocument(ValidateNative);
        }
    }

    private static string FingerprintNative(NativeNode root) {
        var value = new StringBuilder(64 * 1024);
        AppendNative(root, value);
        return Hash(value);
    }

    private static void AppendNative(NativeNode node, StringBuilder value) {
        if (node is NativeElement element) {
            value.Append("E|").Append(element.NamespaceUri).Append('|').Append(element.Prefix).Append('|').Append(element.LocalName);
            foreach (AngleSharp.Dom.IAttr attribute in element.Attributes) {
                value.Append("|A|").Append(attribute.NamespaceUri).Append('|').Append(attribute.Prefix).Append('|')
                    .Append(attribute.LocalName).Append('|').Append(attribute.Value);
            }
        } else if (node is NativeDocumentType type) {
            value.Append("D|").Append(type.Name).Append('|').Append(type.PublicIdentifier).Append('|').Append(type.SystemIdentifier);
        } else if (node.NodeType == AngleSharp.Dom.NodeType.Document) {
            value.Append("R");
        } else if (node.NodeType == AngleSharp.Dom.NodeType.DocumentFragment) {
            value.Append("F");
        } else if (node.NodeType == AngleSharp.Dom.NodeType.Text) {
            value.Append("T|").Append(node.TextContent);
        } else if (node.NodeType == AngleSharp.Dom.NodeType.Comment) {
            value.Append("C|").Append(node.TextContent);
        } else {
            value.Append("N|").Append(node.NodeName).Append('|').Append(node.TextContent);
        }
        value.Append('[');
        foreach (NativeNode child in node.ChildNodes) AppendNative(child, value);
        if (node is NativeTemplateElement template) {
            value.Append("Q[");
            foreach (NativeNode child in template.Content.ChildNodes) AppendNative(child, value);
            value.Append(']');
        }
        value.Append(']');
    }

    private static string FingerprintOwned(HtmlNode root) {
        var value = new StringBuilder(64 * 1024);
        AppendOwned(root, value);
        return Hash(value);
    }

    private static void AppendOwned(HtmlNode node, StringBuilder value) {
        if (node is HtmlElement element) {
            value.Append("E|").Append(element.NamespaceUri).Append('|').Append(element.Prefix).Append('|').Append(element.LocalName);
            foreach (HtmlAttribute attribute in element.Attributes) {
                value.Append("|A|").Append(attribute.NamespaceUri).Append('|').Append(attribute.Prefix).Append('|')
                    .Append(attribute.LocalName).Append('|').Append(attribute.Value);
            }
        } else if (node is HtmlDocumentType type) {
            value.Append("D|").Append(type.Name).Append('|').Append(type.PublicIdentifier).Append('|').Append(type.SystemIdentifier);
        } else if (node.Kind == HtmlNodeKind.Document) {
            value.Append("R");
        } else if (node.Kind == HtmlNodeKind.DocumentFragment) {
            value.Append("F");
        } else if (node.Kind == HtmlNodeKind.Text) {
            value.Append("T|").Append(node.TextContent);
        } else if (node.Kind == HtmlNodeKind.Comment) {
            value.Append("C|").Append(node.TextContent);
        } else {
            value.Append("N|").Append(node.TextContent);
        }
        value.Append('[');
        foreach (HtmlNode child in node.ChildNodes) AppendOwned(child, value);
        if (node is HtmlElement template && template.TemplateContent != null) {
            value.Append("Q[");
            foreach (HtmlNode child in template.TemplateContent.ChildNodes) AppendOwned(child, value);
            value.Append(']');
        }
        value.Append(']');
    }

    private static string Hash(StringBuilder value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value.ToString()))).ToLowerInvariant();

    private readonly record struct ProviderEvidenceValidation(int ResultItems, string FingerprintSha256);
}
