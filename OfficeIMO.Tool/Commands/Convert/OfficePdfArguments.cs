using System.Globalization;

namespace OfficeIMO.Tool.Commands.Convert;

internal sealed class OfficePdfArguments {
    internal const long DefaultMaxInputBytes = 64L * 1024L * 1024L;
    internal const long DefaultMaxOutputBytes = 256L * 1024L * 1024L;
    internal const long DefaultMaxCharactersInPart = 10_000_000L;

    internal bool Help { get; private set; }
    internal string? InputPath { get; private set; }
    internal string? OutputPath { get; private set; }
    internal bool Force { get; private set; }
    internal bool AllowLegacyLoss { get; private set; }
    internal bool RequireNoLoss { get; private set; }
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

    internal static OfficePdfArguments Parse(string[] args) {
        ArgumentNullException.ThrowIfNull(args);
        if (args.Length == 0 || IsHelp(args[0])) return new OfficePdfArguments { Help = true };

        var parsed = new OfficePdfArguments();
        for (int index = 0; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) return new OfficePdfArguments { Help = true };

            switch (token) {
                case "--output":
                case "-o":
                    if (parsed.OutputPath != null) {
                        throw new OfficePdfUsageException("Only one output document may be specified.");
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
                    parsed.RequireNoLoss = true;
                    break;
                case "--diagram-layers":
                    parsed.HasDiagramOptions = true;
                    parsed.DiagramForPrint = NextValue(args, ref index, token) switch {
                        "print" => true,
                        "screen" => false,
                        _ => throw new OfficePdfUsageException("Choose diagram layers: screen or print.")
                    };
                    break;
                case "--text-encoding":
                    parsed.HasTextOptions = true;
                    parsed.TextEncoding = NextValue(args, ref index, token);
                    break;
                case "--tab-size":
                    parsed.HasTextOptions = true;
                    long tabs = ParsePositiveLong(NextValue(args, ref index, token), token);
                    if (tabs > 32) throw new OfficePdfUsageException("Tab size must be between 1 and 32.");
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
                        throw new OfficePdfUsageException("Unknown option '" + token + "'.");
                    }
                    if (parsed.InputPath == null) {
                        parsed.InputPath = token;
                    } else if (parsed.OutputPath == null) {
                        parsed.OutputPath = token;
                    } else {
                        throw new OfficePdfUsageException("Only one input and one output document may be specified.");
                    }
                    break;
            }
        }

        parsed.Validate();
        return parsed;
    }

    private void Validate() {
        if (string.IsNullOrWhiteSpace(InputPath)) {
            throw new OfficePdfUsageException("The convert command requires an input DOC, DOCX, TXT, XLSX, PPTX, ODG, FODG, or DjVu file.");
        }

        string extension = Path.GetExtension(InputPath).ToLowerInvariant();
        if (extension is not ".doc" and not ".docx" and not ".txt" and not ".xlsx" and not ".pptx" and not ".odg" and not ".fodg" and not ".djvu" and not ".djv") {
            throw new OfficePdfUsageException("The convert command supports DOC, DOCX, TXT, XLSX, PPTX, ODG, FODG, and DjVu input.");
        }
        if (AllowLegacyLoss && extension != ".doc") throw new OfficePdfUsageException("--allow-legacy-loss requires DOC input.");
        if (HasTextOptions && extension != ".txt") throw new OfficePdfUsageException("Text encoding and tabs require TXT input.");
        if (HasDiagramOptions && !IsDrawInput) throw new OfficePdfUsageException("Diagram options require ODG or FODG input.");
        if (RequireNoLoss && !IsDrawInput && !IsDjVuInput) throw new OfficePdfUsageException("--require-no-loss requires ODG, FODG, or DjVu input.");

        OutputPath ??= Path.ChangeExtension(InputPath, ".pdf");
        if (Path.GetExtension(OutputPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase) == false) {
            throw new OfficePdfUsageException("The output path must use the .pdf extension.");
        }
        if (OfficeImoToolPathSafety.PathsEqual(InputPath, OutputPath)) {
            throw new OfficePdfUsageException("Input and output paths must be different.");
        }
    }

    private static string NextValue(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index])) {
            throw new OfficePdfUsageException(option + " requires a value.");
        }
        return args[index];
    }

    private static long ParsePositiveLong(string value, string option) {
        if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long parsed) || parsed <= 0) {
            throw new OfficePdfUsageException(option + " requires a positive integer.");
        }
        return parsed;
    }

    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

internal sealed class OfficePdfUsageException : Exception {
    internal OfficePdfUsageException(string message) : base(message) { }
}
