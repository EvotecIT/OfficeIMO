using OfficeIMO.Markup;

namespace OfficeIMO.Tool.Commands.Markup;

internal sealed class MarkupArguments {
    internal const long DefaultMaxInputBytes = 64L * 1024L * 1024L;

    internal string Command { get; private set; } = string.Empty;
    internal string? InputPath { get; private set; }
    internal string? OutputPath { get; private set; }
    internal string? Target { get; private set; }
    internal OfficeMarkupProfile Profile { get; private set; } = OfficeMarkupProfile.Document;
    internal bool UseStdin { get; private set; }
    internal bool ShowHelp { get; private set; }
    internal string? MermaidRendererPath { get; private set; }
    internal bool RenderMermaidDiagrams { get; private set; } = true;
    internal bool WorkbookSafePreflight { get; private set; } = true;
    internal bool WorkbookValidateOpenXml { get; private set; } = true;
    internal bool WorkbookRepairDefinedNames { get; private set; } = true;
    internal long MaxInputBytes { get; private set; } = DefaultMaxInputBytes;
    internal static MarkupArguments Parse(string[] args) {
        var options = new MarkupArguments();
        var positionals = new List<string>();
        for (int index = 0; index < args.Length; index++) {
            string argument = args[index];
            switch (argument) {
                case "-h":
                case "--help":
                    options.ShowHelp = true;
                    break;
                case "--stdin":
                    options.UseStdin = true;
                    break;
                case "--profile":
                    options.Profile = ParseProfile(ReadValue(args, ref index, argument));
                    break;
                case "--target":
                    options.Target = ReadValue(args, ref index, argument);
                    break;
                case "--output":
                case "-o":
                    options.OutputPath = ReadValue(args, ref index, argument);
                    break;
                case "--mermaid-renderer":
                    options.MermaidRendererPath = ReadValue(args, ref index, argument);
                    break;
                case "--no-mermaid":
                    options.RenderMermaidDiagrams = false;
                    break;
                case "--no-safe-preflight":
                    options.WorkbookSafePreflight = false;
                    break;
                case "--no-openxml-validation":
                    options.WorkbookValidateOpenXml = false;
                    break;
                case "--no-defined-name-repair":
                    options.WorkbookRepairDefinedNames = false;
                    break;
                case "--max-input-bytes":
                    string maximumBytes = ReadValue(args, ref index, argument);
                    if (!long.TryParse(maximumBytes, out long parsedMaximumBytes) || parsedMaximumBytes < 1) {
                        throw new ArgumentException("--max-input-bytes must be a positive integer.");
                    }
                    options.MaxInputBytes = parsedMaximumBytes;
                    break;
                case "--format":
                    _ = ReadValue(args, ref index, argument);
                    break;
                default:
                    if (argument.StartsWith("-", StringComparison.Ordinal) && argument != "-") {
                        throw new ArgumentException($"Unknown option '{argument}'.");
                    }
                    positionals.Add(argument);
                    break;
            }
        }

        if (positionals.Count > 0) options.Command = positionals[0].ToLowerInvariant();
        if (positionals.Count > 1) options.InputPath = positionals[1];
        if (positionals.Count > 2) throw new ArgumentException("Only one input path may be specified.");
        if (!options.ShowHelp &&
            options.Command is not ("parse" or "preview" or "validate" or "emit" or "export")) {
            throw new ArgumentException($"Unknown command '{options.Command}'.");
        }

        return options;
    }

    private static string ReadValue(string[] args, ref int index, string option) {
        if (index + 1 >= args.Length) {
            throw new ArgumentException($"Option '{option}' requires a value.");
        }
        return args[++index];
    }

    private static OfficeMarkupProfile ParseProfile(string value) {
        if (Enum.TryParse(value, ignoreCase: true, out OfficeMarkupProfile profile)) return profile;
        throw new ArgumentException($"Unsupported profile '{value}'.");
    }
}
