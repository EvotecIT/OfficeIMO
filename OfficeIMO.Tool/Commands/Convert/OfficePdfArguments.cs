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
                    parsed.OutputPath = NextValue(args, ref index, token);
                    break;
                case "--force":
                    parsed.Force = true;
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
                    if (parsed.InputPath != null) {
                        throw new OfficePdfUsageException("Only one input document may be specified.");
                    }
                    parsed.InputPath = token;
                    break;
            }
        }

        parsed.Validate();
        return parsed;
    }

    private void Validate() {
        if (string.IsNullOrWhiteSpace(InputPath)) {
            throw new OfficePdfUsageException("The convert command requires an input DOCX, XLSX, or PPTX file.");
        }

        string extension = Path.GetExtension(InputPath).ToLowerInvariant();
        if (extension is not ".docx" and not ".xlsx" and not ".pptx") {
            throw new OfficePdfUsageException("The convert command supports DOCX, XLSX, and PPTX input.");
        }

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
