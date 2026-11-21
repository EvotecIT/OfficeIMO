using System.Globalization;
using System.Text;
using System.Text.Json;
using OfficeIMO.Yaml;

namespace OfficeIMO.Tool.Commands.Yaml;

internal static class YamlCommand {
    internal const string Usage = """
OfficeIMO.Tool - YAML lint and format

Usage:
  officeimo yaml lint <path|->... [--format text|json] [--fix] [--strict]
                      [--enable <rule>[,<rule>...]] [--disable <rule>[,<rule>...]]
                      [--max-line-length <n>] [--indent <1-9>]
  officeimo yaml format <path|->... [--write|--check] [--indent <1-9>]
                        [--sequences preserve|compact|indented] [--quotes preserve|single|double]
                        [--line-ending preserve|lf|crlf] [--no-final-newline]

A path may be a file or a folder; folders are searched for *.yaml and *.yml files. '-' reads standard input.
lint prints one finding per line as path:line:column: severity message (rule), or a JSON report with the
schema officeimo.yaml.lint.v1. It returns exit code 1 when a file has errors, or warnings with --strict.
--fix rewrites each file with the safe fixes applied: only the lines a fix repairs change.
Rules: indentation, trailing-spaces, line-length, truthy, key-duplicates, comments, octal-values, tabs,
new-line-at-end-of-file and empty-lines are on by default; document-start, empty-values, quoted-strings
and key-ordering are off. syntax errors are always reported.
format writes one file, or standard input, to standard output; a folder or several files need --write
or --check. --write rewrites the files in place, and --check lists the files that formatting would
change and returns exit code 1 if there are any.
Formatting keeps values, anchors, tags and comments outside flow collections; options left out
follow each file's own conventions.
Files keep their encoding and byte-order mark.
""";

    private sealed class Arguments {
        public bool Lint;
        public readonly List<string> Inputs = new();
        public bool Json;
        public bool Fix;
        public bool Strict;
        public bool Write;
        public bool Check;
        public readonly YamlLintOptions LintOptions = new();
        public readonly YamlFormatOptions FormatOptions = new();
    }

    private class UsageException : Exception {
        public UsageException(string message) : base(message) { }
    }

    internal static async Task<int> RunAsync(string[] args, Stream standardInput, TextWriter standardOutput, TextWriter standardError,
        CancellationToken cancellationToken = default) {
        if (args.Length == 0 || args[0] is "help" || args.Any(static argument => argument is "--help" or "-h")) {
            await standardOutput.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        Arguments parsed;
        List<string> files;
        try {
            parsed = Parse(args);
            files = Expand(parsed.Inputs);
            if (!parsed.Lint && !parsed.Write && !parsed.Check && (files.Count != 1 || parsed.Inputs.Any(Directory.Exists))) {
                throw new UsageException("Formatting a folder or several files needs --write or --check.");
            }
        } catch (UsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (FileNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        }
        try {
            return parsed.Lint
                ? await LintAsync(parsed, files, standardInput, standardOutput, standardError, cancellationToken).ConfigureAwait(false)
                : await FormatAsync(parsed, files, standardInput, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
        } catch (OperationCanceledException) {
            return (int)OfficeImoToolExitCode.Cancelled;
        }
    }

    // ------------------------------------------------------------------ lint

    private static async Task<int> LintAsync(Arguments parsed, List<string> files, Stream standardInput, TextWriter output, TextWriter error, CancellationToken cancellationToken) {
        int errors = 0, warnings = 0, fixedCount = 0;
        bool failedIo = false;
        var results = new List<(string Path, IReadOnlyList<YamlLintFinding> Findings, int Fixed)>();
        foreach (string file in files) {
            cancellationToken.ThrowIfCancellationRequested();
            YamlText source;
            try {
                source = await YamlText.ReadAsync(file, standardInput, cancellationToken).ConfigureAwait(false);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                await error.WriteLineAsync("Could not read " + file + ": " + exception.Message).ConfigureAwait(false);
                failedIo = true;
                continue;
            }
            IReadOnlyList<YamlLintFinding> findings;
            int applied = 0;
            if (parsed.Fix && file != "-") {
                YamlLintFixResult result = YamlLinter.Fix(source.Text, parsed.LintOptions);
                findings = result.Remaining;
                applied = result.Applied.Count;
                if (result.Changed && !await TryWriteAsync(source.With(result.Text), error, cancellationToken).ConfigureAwait(false)) {
                    // The file kept its old text, so report what it still contains.
                    failedIo = true;
                    findings = YamlLinter.Lint(source.Text, parsed.LintOptions);
                    applied = 0;
                }
            } else {
                findings = YamlLinter.Lint(source.Text, parsed.LintOptions);
            }
            fixedCount += applied;
            foreach (YamlLintFinding finding in findings) {
                if (finding.Severity == OfficeDiagnosticSeverity.Error) errors++;
                else if (finding.Severity == OfficeDiagnosticSeverity.Warning) warnings++;
            }
            results.Add((file, findings, applied));
        }
        if (parsed.Json) {
            await WriteLintJsonAsync(output, results, errors, warnings, fixedCount).ConfigureAwait(false);
        } else {
            foreach ((string path, IReadOnlyList<YamlLintFinding> findings, _) in results) {
                foreach (YamlLintFinding finding in findings) {
                    await output.WriteLineAsync(Display(path) + ":" + finding.Line + ":" + finding.Column + ": " + SeverityName(finding.Severity) + " " + finding.Message + " (" + finding.RuleId + ")").ConfigureAwait(false);
                }
            }
            if (parsed.Fix && fixedCount > 0) await error.WriteLineAsync("Fixed " + fixedCount + " finding(s).").ConfigureAwait(false);
        }
        if (failedIo) return (int)OfficeImoToolExitCode.OutputFailed;
        return errors > 0 || (parsed.Strict && warnings > 0) ? (int)OfficeImoToolExitCode.ValidationFailed : (int)OfficeImoToolExitCode.Success;
    }

    private static async Task WriteLintJsonAsync(TextWriter output, List<(string Path, IReadOnlyList<YamlLintFinding> Findings, int Fixed)> results, int errors, int warnings, int fixedCount) {
        using var bytes = new MemoryStream();
        using (var json = new Utf8JsonWriter(bytes, new JsonWriterOptions { Indented = true })) {
            json.WriteStartObject();
            json.WriteString("schema", "officeimo.yaml.lint.v1");
            json.WriteNumber("errors", errors);
            json.WriteNumber("warnings", warnings);
            json.WriteNumber("fixed", fixedCount);
            json.WriteStartArray("files");
            foreach ((string path, IReadOnlyList<YamlLintFinding> findings, int applied) in results) {
                json.WriteStartObject();
                json.WriteString("path", Display(path));
                json.WriteNumber("fixed", applied);
                json.WriteStartArray("findings");
                foreach (YamlLintFinding finding in findings) {
                    json.WriteStartObject();
                    json.WriteString("rule", finding.RuleId);
                    json.WriteString("severity", SeverityName(finding.Severity));
                    json.WriteNumber("line", finding.Line);
                    json.WriteNumber("column", finding.Column);
                    json.WriteString("message", finding.Message);
                    json.WriteBoolean("fixable", finding.Fix is not null);
                    json.WriteEndObject();
                }
                json.WriteEndArray();
                json.WriteEndObject();
            }
            json.WriteEndArray();
            json.WriteEndObject();
        }
        await output.WriteLineAsync(Encoding.UTF8.GetString(bytes.ToArray())).ConfigureAwait(false);
    }

    private static string SeverityName(OfficeDiagnosticSeverity severity) => severity switch {
        OfficeDiagnosticSeverity.Error => "error",
        OfficeDiagnosticSeverity.Warning => "warning",
        _ => "info"
    };

    // ------------------------------------------------------------------ format

    private static async Task<int> FormatAsync(Arguments parsed, List<string> files, Stream standardInput, TextWriter output, TextWriter error, CancellationToken cancellationToken) {
        int changed = 0;
        bool failedIo = false, invalid = false;
        foreach (string file in files) {
            cancellationToken.ThrowIfCancellationRequested();
            YamlText source;
            try {
                source = await YamlText.ReadAsync(file, standardInput, cancellationToken).ConfigureAwait(false);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                await error.WriteLineAsync("Could not read " + file + ": " + exception.Message).ConfigureAwait(false);
                failedIo = true;
                continue;
            }
            string formatted;
            try {
                formatted = YamlFormatter.Format(source.Text, parsed.FormatOptions);
            } catch (YamlException exception) {
                await error.WriteLineAsync(Display(file) + ":" + exception.Line + ":" + exception.Column + ": error " + FirstLine(exception.Message) + " (syntax)").ConfigureAwait(false);
                invalid = true;
                continue;
            }
            bool differs = !string.Equals(formatted, source.Text, StringComparison.Ordinal);
            if (parsed.Check) {
                if (differs) {
                    changed++;
                    await output.WriteLineAsync(Display(file)).ConfigureAwait(false);
                }
            } else if (parsed.Write && file != "-") {
                if (differs) {
                    changed++;
                    if (!await TryWriteAsync(source.With(formatted), error, cancellationToken).ConfigureAwait(false)) failedIo = true;
                }
            } else {
                await output.WriteAsync(formatted).ConfigureAwait(false);
            }
        }
        if (failedIo) return (int)OfficeImoToolExitCode.OutputFailed;
        if (invalid) return (int)OfficeImoToolExitCode.UnsupportedInput;
        return parsed.Check && changed > 0 ? (int)OfficeImoToolExitCode.ValidationFailed : (int)OfficeImoToolExitCode.Success;
    }

    private static string FirstLine(string message) {
        int end = message.IndexOfAny(['\r', '\n']);
        return end < 0 ? message : message.Substring(0, end);
    }

    // ------------------------------------------------------------------ files

    private static async Task<bool> TryWriteAsync(YamlText text, TextWriter error, CancellationToken cancellationToken) {
        try {
            await text.WriteAsync(cancellationToken).ConfigureAwait(false);
            return true;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            await error.WriteLineAsync("Could not write " + text.Path + ": " + exception.Message).ConfigureAwait(false);
            return false;
        }
    }

    private static string Display(string path) => path == "-" ? "<stdin>" : path;

    private static List<string> Expand(List<string> inputs) {
        var files = new List<string>();
        foreach (string input in inputs) {
            if (input == "-") {
                files.Add(input);
            } else if (Directory.Exists(input)) {
                files.AddRange(Directory.EnumerateFiles(input, "*.*", SearchOption.AllDirectories)
                    .Where(static path => path.EndsWith(".yaml", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".yml", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(static path => path, StringComparer.Ordinal));
            } else if (File.Exists(input)) {
                files.Add(input);
            } else {
                throw new FileNotFoundException("Input was not found: " + input);
            }
        }
        if (files.Count(static file => file == "-") > 1) throw new UsageException("Standard input can be read once.");
        return files;
    }

    // ------------------------------------------------------------------ arguments

    private static Arguments Parse(string[] args) {
        var parsed = new Arguments {
            Lint = args[0] switch {
                "lint" => true,
                "format" => false,
                _ => throw new UsageException("Unknown yaml command '" + args[0] + "'.")
            }
        };
        for (int i = 1; i < args.Length; i++) {
            string argument = args[i];
            if (argument == "-" || !argument.StartsWith("--", StringComparison.Ordinal)) {
                parsed.Inputs.Add(argument);
                continue;
            }
            switch (argument) {
                case "--format" when parsed.Lint:
                    parsed.Json = Value(args, ref i) switch {
                        "json" => true,
                        "text" => false,
                        _ => throw new UsageException("--format must be text or json.")
                    };
                    break;
                case "--fix" when parsed.Lint:
                    parsed.Fix = true;
                    break;
                case "--strict" when parsed.Lint:
                    parsed.Strict = true;
                    break;
                case "--enable" when parsed.Lint:
                case "--disable" when parsed.Lint:
                    foreach (string id in Value(args, ref i).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) {
                        YamlLintRuleOptions rule = parsed.LintOptions.GetRule(id) ?? throw new UsageException("Unknown lint rule '" + id + "'.");
                        rule.Enabled = argument == "--enable";
                    }
                    break;
                case "--max-line-length" when parsed.Lint:
                    parsed.LintOptions.LineLength.Max = Number(argument, Value(args, ref i), 1, 100000);
                    break;
                case "--indent":
                    int indent = Number(argument, Value(args, ref i), 1, 9);
                    parsed.LintOptions.Indentation.Spaces = indent;
                    parsed.FormatOptions.IndentWidth = indent;
                    break;
                case "--write" when !parsed.Lint:
                    parsed.Write = true;
                    break;
                case "--check" when !parsed.Lint:
                    parsed.Check = true;
                    break;
                case "--sequences" when !parsed.Lint:
                    parsed.FormatOptions.SequenceIndentation = Value(args, ref i) switch {
                        "preserve" => YamlSequenceIndentation.Preserve,
                        "compact" => YamlSequenceIndentation.Compact,
                        "indented" => YamlSequenceIndentation.Indented,
                        _ => throw new UsageException("--sequences must be preserve, compact or indented.")
                    };
                    break;
                case "--quotes" when !parsed.Lint:
                    parsed.FormatOptions.Quotes = Value(args, ref i) switch {
                        "preserve" => YamlQuoteStyle.Preserve,
                        "single" => YamlQuoteStyle.Single,
                        "double" => YamlQuoteStyle.Double,
                        _ => throw new UsageException("--quotes must be preserve, single or double.")
                    };
                    break;
                case "--line-ending" when !parsed.Lint:
                    parsed.FormatOptions.LineEnding = Value(args, ref i) switch {
                        "preserve" => YamlLineEnding.Preserve,
                        "lf" => YamlLineEnding.Lf,
                        "crlf" => YamlLineEnding.CrLf,
                        _ => throw new UsageException("--line-ending must be preserve, lf or crlf.")
                    };
                    break;
                case "--no-final-newline" when !parsed.Lint:
                    parsed.FormatOptions.FinalNewline = false;
                    break;
                default:
                    throw new UsageException("Unknown option '" + argument + "' for yaml " + args[0] + ".");
            }
        }
        if (parsed.Inputs.Count == 0) throw new UsageException("Name at least one file, folder or '-'.");
        if (parsed.Write && parsed.Check) throw new UsageException("--write and --check cannot be combined.");
        if ((parsed.Write || parsed.Fix) && parsed.Inputs.Contains("-")) throw new UsageException("Standard input cannot be rewritten; leave out --write and --fix.");
        return parsed;
    }

    private static string Value(string[] args, ref int i) {
        if (i + 1 >= args.Length) throw new UsageException(args[i] + " needs a value.");
        return args[++i];
    }

    private static int Number(string option, string text, int min, int max) =>
        int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value) && value >= min && value <= max
            ? value
            : throw new UsageException(option + " must be a number from " + min + " to " + max + ".");
}

/// <summary>A YAML file's text with the encoding and byte-order mark it was read with, so a rewrite keeps both.</summary>
internal sealed class YamlText {
    private readonly Encoding _encoding;
    private readonly bool _byteOrderMark;

    private YamlText(string path, string text, Encoding encoding, bool byteOrderMark) {
        Path = path;
        Text = text;
        _encoding = encoding;
        _byteOrderMark = byteOrderMark;
    }

    public string Path { get; }
    public string Text { get; }

    public YamlText With(string text) => new(Path, text, _encoding, _byteOrderMark);

    public static async Task<YamlText> ReadAsync(string path, Stream standardInput, CancellationToken cancellationToken) {
        byte[] bytes;
        if (path == "-") {
            using var buffer = new MemoryStream();
            await standardInput.CopyToAsync(buffer, cancellationToken).ConfigureAwait(false);
            bytes = buffer.ToArray();
        } else {
            bytes = await File.ReadAllBytesAsync(path, cancellationToken).ConfigureAwait(false);
        }
        using var reader = new StreamReader(new MemoryStream(bytes), new UTF8Encoding(false), detectEncodingFromByteOrderMarks: true);
        string text = await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        Encoding encoding = reader.CurrentEncoding;
        byte[] preamble = encoding.GetPreamble();
        bool mark = preamble.Length > 0 && bytes.AsSpan().StartsWith(preamble);
        return new YamlText(path, text, encoding, mark);
    }

    /// <summary>Replaces the file through a temporary sibling, so a failed write leaves the old contents.</summary>
    public async Task WriteAsync(CancellationToken cancellationToken) {
        string full = System.IO.Path.GetFullPath(Path);
        string temporary = full + ".officeimo-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".tmp";
        try {
            var streamOptions = new FileStreamOptions { Mode = FileMode.CreateNew, Access = FileAccess.Write, Share = FileShare.None };
            UnixFileMode? mode = null;
            if (!OperatingSystem.IsWindows()) {
                // The replacement keeps the file's permissions, and is never readable by more users while it is written.
                mode = File.GetUnixFileMode(full);
                streamOptions.UnixCreateMode = mode.Value;
            }
            await using (var stream = new FileStream(temporary, streamOptions)) {
                if (_byteOrderMark) {
                    byte[] preamble = _encoding.GetPreamble();
                    await stream.WriteAsync(preamble, cancellationToken).ConfigureAwait(false);
                }
                byte[] content = _encoding.GetBytes(Text);
                await stream.WriteAsync(content, cancellationToken).ConfigureAwait(false);
            }
            if (mode is { } unixMode && !OperatingSystem.IsWindows()) File.SetUnixFileMode(temporary, unixMode);
            File.Move(temporary, full, overwrite: true);
        } finally {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }
}
