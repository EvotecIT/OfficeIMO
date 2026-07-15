using System.Globalization;

namespace OfficeIMO.Reader.Tool;

internal enum ReaderToolCommand {
    Help,
    Read,
    Folder,
    Capabilities
}

internal enum ReaderToolOutputFormat {
    Markdown,
    Json,
    Text
}

internal sealed class ReaderToolArguments {
    internal const long DefaultMaxInputBytes = 64L * 1024L * 1024L;

    internal ReaderToolCommand Command { get; private set; }
    internal ReaderToolOutputFormat Format { get; private set; }
    internal string? InputPath { get; private set; }
    internal string? SourceName { get; private set; }
    internal string? OutputPath { get; private set; }
    internal string? AssetsPath { get; private set; }
    internal int Concurrency { get; private set; } = 4;
    internal int MaxFiles { get; private set; } = 500;
    internal long? MaxTotalBytes { get; private set; }
    internal long MaxInputBytes { get; private set; } = DefaultMaxInputBytes;
    internal bool Recurse { get; private set; } = true;
    private bool ConcurrencySpecified { get; set; }
    private bool FolderLimitSpecified { get; set; }
    private bool ReadLimitSpecified { get; set; }
    private bool RecursionSpecified { get; set; }

    internal static ReaderToolArguments Parse(string[] args) {
        if (args == null) throw new ArgumentNullException(nameof(args));
        if (args.Length == 0 || IsHelp(args[0])) {
            return new ReaderToolArguments { Command = ReaderToolCommand.Help };
        }

        var parsed = new ReaderToolArguments {
            Command = ParseCommand(args[0])
        };
        parsed.Format = parsed.Command == ReaderToolCommand.Capabilities
            ? ReaderToolOutputFormat.Text
            : ReaderToolOutputFormat.Markdown;

        for (int index = 1; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) {
                parsed.Command = ReaderToolCommand.Help;
                return parsed;
            }

            switch (token) {
                case "--format":
                    parsed.Format = ParseFormat(NextValue(args, ref index, token));
                    break;
                case "--output":
                case "-o":
                    parsed.OutputPath = NextValue(args, ref index, token);
                    break;
                case "--assets":
                    parsed.AssetsPath = NextValue(args, ref index, token);
                    break;
                case "--name":
                    parsed.SourceName = NextValue(args, ref index, token);
                    break;
                case "--concurrency":
                    parsed.Concurrency = ParseBoundedInt(NextValue(args, ref index, token), token, 1, 64);
                    parsed.ConcurrencySpecified = true;
                    break;
                case "--max-files":
                    parsed.MaxFiles = ParseBoundedInt(NextValue(args, ref index, token), token, 1, 100_000);
                    parsed.FolderLimitSpecified = true;
                    break;
                case "--max-total-bytes":
                    parsed.MaxTotalBytes = ParsePositiveLong(NextValue(args, ref index, token), token);
                    parsed.FolderLimitSpecified = true;
                    break;
                case "--max-input-bytes":
                    parsed.MaxInputBytes = ParsePositiveLong(NextValue(args, ref index, token), token);
                    parsed.ReadLimitSpecified = true;
                    break;
                case "--recursive":
                    parsed.Recurse = true;
                    parsed.RecursionSpecified = true;
                    break;
                case "--no-recursive":
                    parsed.Recurse = false;
                    parsed.RecursionSpecified = true;
                    break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal) && token != "-") {
                        throw new ReaderToolUsageException("Unknown option '" + token + "'.");
                    }
                    if (parsed.InputPath != null) {
                        throw new ReaderToolUsageException("Only one input path may be specified.");
                    }
                    parsed.InputPath = token;
                    break;
            }
        }

        parsed.Validate();
        return parsed;
    }

    private void Validate() {
        if (Command == ReaderToolCommand.Capabilities) {
            if (InputPath != null || SourceName != null || AssetsPath != null || OutputPath != null ||
                ConcurrencySpecified || FolderLimitSpecified || ReadLimitSpecified || RecursionSpecified) {
                throw new ReaderToolUsageException("The capabilities command does not accept input or output paths.");
            }
            if (Format == ReaderToolOutputFormat.Markdown) {
                throw new ReaderToolUsageException("Capabilities format must be 'text' or 'json'.");
            }
            return;
        }

        if (string.IsNullOrWhiteSpace(InputPath)) {
            throw new ReaderToolUsageException("The " + Command.ToString().ToLowerInvariant() + " command requires an input path.");
        }
        if (Format == ReaderToolOutputFormat.Text) {
            throw new ReaderToolUsageException("Document format must be 'markdown' or 'json'.");
        }

        if (Command == ReaderToolCommand.Read) {
            if (InputPath == "-" && string.IsNullOrWhiteSpace(SourceName)) {
                SourceName = "stdin.txt";
            } else if (InputPath != "-" && SourceName != null) {
                throw new ReaderToolUsageException("--name is only valid when reading standard input.");
            }
            if (ConcurrencySpecified || FolderLimitSpecified || RecursionSpecified) {
                throw new ReaderToolUsageException("Folder traversal options are only valid with the folder command.");
            }
            return;
        }

        if (InputPath == "-") {
            throw new ReaderToolUsageException("The folder command does not read from standard input.");
        }
        if (string.IsNullOrWhiteSpace(OutputPath) || OutputPath == "-") {
            throw new ReaderToolUsageException("The folder command requires --output <directory>.");
        }
        if (SourceName != null) {
            throw new ReaderToolUsageException("--name is only valid when reading standard input.");
        }
        if (ReadLimitSpecified) {
            throw new ReaderToolUsageException("--max-input-bytes is only valid with the read command.");
        }
    }

    private static ReaderToolCommand ParseCommand(string value) {
        return value.ToLowerInvariant() switch {
            "read" => ReaderToolCommand.Read,
            "folder" => ReaderToolCommand.Folder,
            "capabilities" => ReaderToolCommand.Capabilities,
            _ => throw new ReaderToolUsageException("Unknown command '" + value + "'.")
        };
    }

    private static ReaderToolOutputFormat ParseFormat(string value) {
        return value.ToLowerInvariant() switch {
            "markdown" or "md" => ReaderToolOutputFormat.Markdown,
            "json" => ReaderToolOutputFormat.Json,
            "text" => ReaderToolOutputFormat.Text,
            _ => throw new ReaderToolUsageException("Unknown format '" + value + "'.")
        };
    }

    private static string NextValue(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index])) {
            throw new ReaderToolUsageException(option + " requires a value.");
        }
        return args[index];
    }

    private static int ParseBoundedInt(string value, string option, int minimum, int maximum) {
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ||
            parsed < minimum || parsed > maximum) {
            throw new ReaderToolUsageException(option + " must be between " + minimum + " and " + maximum + ".");
        }
        return parsed;
    }

    private static long ParsePositiveLong(string value, string option) {
        if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long parsed) || parsed < 1) {
            throw new ReaderToolUsageException(option + " must be a positive integer.");
        }
        return parsed;
    }

    private static bool IsHelp(string value) {
        return value is "--help" or "-h" or "help";
    }
}

internal sealed class ReaderToolUsageException : Exception {
    internal ReaderToolUsageException(string message) : base(message) { }
}
