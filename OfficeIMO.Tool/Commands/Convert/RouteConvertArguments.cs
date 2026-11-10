using System.Globalization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Convert;

internal sealed class RouteConvertArguments {
    internal const long DefaultMaxInputBytes = 64L * 1024L * 1024L;
    internal const long DefaultMaxOutputBytes = 256L * 1024L * 1024L;
    internal const long DefaultMaxCharactersInPart = 10_000_000L;

    internal bool Help { get; private set; }
    internal string? InputPath { get; private set; }
    internal string? OutputPath { get; private set; }
    internal bool Force { get; private set; }
    internal bool AllowLegacyLoss { get; private set; }
    internal bool FailOnLoss { get; private set; }
    internal bool TextOnly { get; private set; }
    internal OfficeWorkflowRoute? Route { get; private set; }
    internal bool DiagramForPrint { get; private set; } = true;
    private bool HasDiagramOptions { get; set; }
    internal bool IsDrawInput => Path.GetExtension(InputPath ?? string.Empty).ToLowerInvariant() is ".odg" or ".fodg";
    internal bool IsDjVuInput => Path.GetExtension(InputPath ?? string.Empty).ToLowerInvariant() is ".djvu" or ".djv";
    internal string? TextEncoding { get; private set; }
    internal int TabSize { get; private set; } = 8;
    private bool HasTextOptions { get; set; }
    internal long MaxInputBytes { get; private set; } = DefaultMaxInputBytes;
    internal long MaxOutputBytes { get; private set; } = DefaultMaxOutputBytes;
    internal long MaxCharactersInPart { get; private set; } = DefaultMaxCharactersInPart;

    internal static RouteConvertArguments Parse(string[] args) {
        ArgumentNullException.ThrowIfNull(args);
        if (args.Length == 0 || IsHelp(args[0])) return new RouteConvertArguments { Help = true };

        var parsed = new RouteConvertArguments();
        for (int index = 0; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) return new RouteConvertArguments { Help = true };

            switch (token) {
                case "--output":
                case "-o":
                    if (parsed.OutputPath != null) {
                        throw new RouteConvertUsageException("Only one output document may be specified.");
                    }
                    parsed.OutputPath = NextValue(args, ref index, token);
                    break;
                case "--force":
                    parsed.Force = true;
                    break;
                case "--allow-legacy-loss":
                    parsed.AllowLegacyLoss = true;
                    break;
                case "--require-no-loss":
                    parsed.FailOnLoss = true;
                    break;
                case "--text-only":
                    parsed.TextOnly = true;
                    break;
                case "--diagram-layers":
                    parsed.HasDiagramOptions = true;
                    parsed.DiagramForPrint = NextValue(args, ref index, token) switch {
                        "print" => true,
                        "screen" => false,
                        _ => throw new RouteConvertUsageException("Choose diagram layers: screen or print.")
                    };
                    break;
                case "--text-encoding":
                    parsed.HasTextOptions = true;
                    parsed.TextEncoding = NextValue(args, ref index, token);
                    break;
                case "--tab-size":
                    parsed.HasTextOptions = true;
                    long tabs = ParsePositiveLong(NextValue(args, ref index, token), token);
                    if (tabs > 32) throw new RouteConvertUsageException("Tab size must be between 1 and 32.");
                    parsed.TabSize = (int)tabs;
                    break;
                case "--max-input-bytes":
                    parsed.MaxInputBytes = ParsePositiveLong(NextValue(args, ref index, token), token);
                    break;
                case "--max-output-bytes":
                    parsed.MaxOutputBytes = ParsePositiveLong(NextValue(args, ref index, token), token);
                    break;
                case "--max-characters-in-part":
                    parsed.MaxCharactersInPart = ParsePositiveLong(NextValue(args, ref index, token), token);
                    break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal)) {
                        throw new RouteConvertUsageException("Unknown option '" + token + "'.");
                    }
                    if (parsed.InputPath == null) {
                        parsed.InputPath = token;
                    } else if (parsed.OutputPath == null) {
                        parsed.OutputPath = token;
                    } else {
                        throw new RouteConvertUsageException("Only one input and one output document may be specified.");
                    }
                    break;
            }
        }

        parsed.Validate();
        return parsed;
    }

    private void Validate() {
        if (string.IsNullOrWhiteSpace(InputPath)) {
            throw new RouteConvertUsageException("The convert command requires an input document.");
        }

        string extension = Path.GetExtension(InputPath).ToLowerInvariant();
        OutputPath ??= Path.ChangeExtension(InputPath, ".pdf");
        string target = Path.GetExtension(OutputPath).ToLowerInvariant();
        Route = extension.Length == 0 || target.Length == 0 ? null : OfficeWorkflowCatalog.Find(extension, target, executableOnly: true);
        if (Route is null) {
            throw new RouteConvertUsageException("No conversion route converts '" + (extension.Length == 0 ? "(none)" : extension) + "' to '"
                + (target.Length == 0 ? "(none)" : target) + "'. Run 'officeimo convert --help' for the supported routes.");
        }
        if (AllowLegacyLoss && extension != ".doc") throw new RouteConvertUsageException("--allow-legacy-loss requires DOC input.");
        if (HasTextOptions && Route.Id != "txt-pdf") throw new RouteConvertUsageException("Text encoding and tabs require TXT-to-PDF conversion.");
        if (HasDiagramOptions && !IsDrawInput) throw new RouteConvertUsageException("Diagram options require ODG or FODG input.");
        if (FailOnLoss && Route.Id is "doc-pdf" or "docx-pdf" or "txt-pdf" or "xlsx-pdf" or "pptx-pdf")
            throw new RouteConvertUsageException("--require-no-loss is not available for DOC, DOCX, TXT, XLSX or PPTX to PDF.");
        if (TextOnly && Route.Id != "email-pdf") throw new RouteConvertUsageException("--text-only requires email-to-PDF conversion.");
        if (OfficeImoToolPathSafety.PathsEqual(InputPath, OutputPath)) {
            throw new RouteConvertUsageException("Input and output paths must be different.");
        }
    }

    private static string NextValue(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index])) {
            throw new RouteConvertUsageException(option + " requires a value.");
        }
        return args[index];
    }

    private static long ParsePositiveLong(string value, string option) {
        if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long parsed) || parsed <= 0) {
            throw new RouteConvertUsageException(option + " requires a positive integer.");
        }
        return parsed;
    }

    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

internal sealed class RouteConvertUsageException : Exception {
    internal RouteConvertUsageException(string message) : base(message) { }
}
