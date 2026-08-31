namespace OfficeIMO.Tool.Commands.Tabular;

internal enum TabularCommandKind {
    Help,
    Sheets,
    Schema,
    Convert
}

internal sealed class TabularArguments {
    private TabularArguments() { }

    internal TabularCommandKind Command { get; private init; }
    internal string InputPath { get; private init; } = string.Empty;
    internal string? OutputPath { get; private init; }
    internal string? SheetName { get; private init; }
    internal int? SheetIndex { get; private init; }
    internal char? Delimiter { get; private init; }
    internal bool HasHeaderRow { get; private init; } = true;
    internal bool Force { get; private init; }

    internal static TabularArguments Parse(string[] args) {
        ArgumentNullException.ThrowIfNull(args);
        if (args.Length == 0 || IsHelp(args[0])) {
            return new TabularArguments { Command = TabularCommandKind.Help };
        }

        TabularCommandKind command = args[0].ToLowerInvariant() switch {
            "sheets" => TabularCommandKind.Sheets,
            "schema" => TabularCommandKind.Schema,
            "convert" => TabularCommandKind.Convert,
            _ => throw new TabularUsageException("Unknown tabular command '" + args[0] + "'.")
        };
        string? inputPath = null;
        string? outputPath = null;
        string? sheetName = null;
        int? sheetIndex = null;
        char? delimiter = null;
        bool hasHeaderRow = true;
        bool force = false;

        for (int index = 1; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) return new TabularArguments { Command = TabularCommandKind.Help };
            switch (token) {
                case "--sheet":
                    sheetName = NextValue(args, ref index, token);
                    break;
                case "--sheet-index":
                    string indexText = NextValue(args, ref index, token);
                    if (!int.TryParse(indexText, out int parsedIndex) || parsedIndex < 0) {
                        throw new TabularUsageException("--sheet-index requires a non-negative zero-based index.");
                    }
                    sheetIndex = parsedIndex;
                    break;
                case "--delimiter":
                    string delimiterText = NextValue(args, ref index, token);
                    delimiter = delimiterText == "\\t"
                        ? '\t'
                        : delimiterText.Length == 1
                            ? delimiterText[0]
                            : throw new TabularUsageException("--delimiter requires one character or \\t.");
                    break;
                case "--no-header":
                    hasHeaderRow = false;
                    break;
                case "--force":
                    force = true;
                    break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal)) {
                        throw new TabularUsageException("Unknown option '" + token + "'.");
                    }
                    if (inputPath == null) inputPath = token;
                    else if (outputPath == null) outputPath = token;
                    else throw new TabularUsageException("Only one input and one output path may be specified.");
                    break;
            }
        }

        if (string.IsNullOrWhiteSpace(inputPath)) {
            throw new TabularUsageException(command.ToString().ToLowerInvariant() + " requires an input path.");
        }
        if (sheetName != null && sheetIndex.HasValue) {
            throw new TabularUsageException("--sheet and --sheet-index cannot be combined.");
        }
        if (command == TabularCommandKind.Convert && string.IsNullOrWhiteSpace(outputPath)) {
            throw new TabularUsageException("convert requires an output path.");
        }
        if (command != TabularCommandKind.Convert && outputPath != null) {
            throw new TabularUsageException(command.ToString().ToLowerInvariant() + " accepts only one input path.");
        }
        if (command == TabularCommandKind.Sheets &&
            (sheetName != null || sheetIndex.HasValue || delimiter.HasValue || !hasHeaderRow || force)) {
            throw new TabularUsageException("sheets does not accept selection, delimiter, header, or overwrite options.");
        }

        return new TabularArguments {
            Command = command,
            InputPath = Path.GetFullPath(inputPath),
            OutputPath = outputPath == null ? null : Path.GetFullPath(outputPath),
            SheetName = sheetName,
            SheetIndex = sheetIndex,
            Delimiter = delimiter,
            HasHeaderRow = hasHeaderRow,
            Force = force
        };
    }

    private static string NextValue(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index])) {
            throw new TabularUsageException(option + " requires a value.");
        }
        return args[index];
    }

    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

internal sealed class TabularUsageException : Exception {
    internal TabularUsageException(string message) : base(message) { }
}
