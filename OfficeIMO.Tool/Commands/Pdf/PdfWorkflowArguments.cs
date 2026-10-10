using System.Globalization;
using OfficeIMO.Tool.Agent;

namespace OfficeIMO.Tool.Commands.Pdf;

internal sealed class PdfWorkflowArguments {
    internal string Operation { get; private set; } = "";
    internal string? Input { get; private set; }
    internal string? Output { get; private set; }
    internal string? Provider { get; private set; }
    internal string? Language { get; private set; }
    internal double Confidence { get; private set; } = 0.5;
    internal int PagesPerDocument { get; private set; } = 1;
    internal bool Help { get; private set; }
    internal IList<string> ProviderAssemblies { get; } = new List<string>();
    internal Dictionary<string, string> ProviderOptions { get; } = new(StringComparer.Ordinal);
    internal PdfWorkflowSettings Settings { get; private set; } = new();

    internal static PdfWorkflowArguments Parse(string[] args) {
        var result = new PdfWorkflowArguments();
        if (args.Length == 0 || IsHelp(args[0])) { result.Help = true; return result; }
        result.Operation = args[0].ToLowerInvariant();
        if (result.Operation is not ("extract" or "split" or "decrypt" or "flatten" or "ocr" or "optimize" or "sanitize" or "providers"))
            throw new AgentUsageException("Unknown PDF workflow '" + args[0] + "'.");
        string? pages = null, password = null;
        bool force = false, acknowledge = false;
        int maximumPages = 100;
        double dpi = 150;
        long maximumPixels = 25_000_000, inputBytes = 256L * 1024 * 1024, outputBytes = 512L * 1024 * 1024;
        for (int index = 1; index < args.Length; index++) {
            string option = args[index];
            if (IsHelp(option)) { result.Help = true; return result; }
            switch (option) {
                case "--output": result.Output = Value(args, ref index, option); break;
                case "--pages": Only(result.Operation, option, "extract", "flatten", "ocr"); pages = Value(args, ref index, option); break;
                case "--pages-per-document": Only(result.Operation, option, "split"); result.PagesPerDocument = Integer(Value(args, ref index, option), option); break;
                case "--maximum-pages": maximumPages = Integer(Value(args, ref index, option), option); break;
                case "--dpi": Only(result.Operation, option, "flatten", "ocr"); dpi = Number(Value(args, ref index, option), option); break;
                case "--maximum-pixels-per-page": Only(result.Operation, option, "flatten", "ocr"); maximumPixels = Long(Value(args, ref index, option), option); break;
                case "--maximum-input-bytes": inputBytes = Long(Value(args, ref index, option), option); break;
                case "--maximum-output-bytes": outputBytes = Long(Value(args, ref index, option), option); break;
                case "--password-env": password = Value(args, ref index, option); break;
                case "--force": force = true; break;
                case "--acknowledge-raster-output": Only(result.Operation, option, "flatten"); acknowledge = true; break;
                case "--ocr-provider": Only(result.Operation, option, "ocr"); result.Provider = Value(args, ref index, option); break;
                case "--ocr-language": Only(result.Operation, option, "ocr"); result.Language = Value(args, ref index, option); break;
                case "--ocr-min-confidence": Only(result.Operation, option, "ocr"); result.Confidence = Number(Value(args, ref index, option), option); break;
                case "--ocr-provider-assembly": Only(result.Operation, option, "ocr", "providers"); result.ProviderAssemblies.Add(Value(args, ref index, option)); break;
                case "--ocr-option": Only(result.Operation, option, "ocr"); AddProviderOption(result.ProviderOptions, Value(args, ref index, option)); break;
                default:
                    if (option.StartsWith('-')) throw new AgentUsageException("Unknown PDF option '" + option + "'.");
                    if (result.Input is not null) throw new AgentUsageException("PDF workflows accept one input file.");
                    result.Input = option; break;
            }
        }
        result.Settings = new PdfWorkflowSettings {
            Pages = pages, PasswordEnvironmentVariable = password, Overwrite = force, AcknowledgeRasterOutput = acknowledge,
            MaximumPages = maximumPages, Dpi = dpi, MaximumPixelsPerPage = maximumPixels,
            MaximumInputBytes = inputBytes, MaximumOutputBytes = outputBytes
        };
        result.Settings.Validate();
        if (result.Operation == "providers") {
            if (args.Skip(1).Where((_, index) => index % 2 == 0).Any(token => token != "--ocr-provider-assembly"))
                throw new AgentUsageException("providers accepts only --ocr-provider-assembly options.");
            return result;
        }
        if (result.Input is null || result.Output is null) throw new AgentUsageException("PDF workflows require one input and --output <new-path>.");
        if (result.Operation == "ocr" && result.Provider is null) throw new AgentUsageException("ocr requires --ocr-provider <registered-id>.");
        return result;
    }

    internal static void AddProviderOption(Dictionary<string, string> options, string value) {
        int split = value.IndexOf('=');
        if (split <= 0 || !options.TryAdd(value[..split], value[(split + 1)..]))
            throw new AgentUsageException("OCR options require unique key=value entries.");
        if (options.Count > 128 || options.Sum(pair => (long)pair.Key.Length + pair.Value.Length) > 64 * 1024)
            throw new AgentUsageException("OCR options exceed their scalar configuration budget.");
    }

    internal static string Value(string[] args, ref int index, string option) => ++index < args.Length &&
        !string.IsNullOrWhiteSpace(args[index]) && !args[index].StartsWith('-') ? args[index] : throw new AgentUsageException(option + " requires a value.");
    private static int Integer(string value, string option) => int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int number) ? number : throw new AgentUsageException(option + " requires an integer.");
    private static long Long(string value, string option) => long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long number) ? number : throw new AgentUsageException(option + " requires an integer.");
    private static double Number(string value, string option) => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) && double.IsFinite(number) ? number : throw new AgentUsageException(option + " requires a finite number.");
    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
    private static void Only(string operation, string option, params string[] allowed) {
        if (!allowed.Contains(operation)) throw new AgentUsageException(option + " is not supported by " + operation + ".");
    }
}
