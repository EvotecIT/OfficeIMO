using System.Globalization;
using System.Text;

namespace OfficeIMO.Tool.Commands.Accessibility;

internal static class AccessibilityCommand {
    internal const string Usage = """
OfficeIMO.Tool - accessibility inspection

Usage:
  officeimo accessibility <input> [--format json|sarif|text] [--strict] [--fail-on-warnings]
             [--output <path>] [--password <password>]

Inputs are Word, Excel and PowerPoint files. All three return the same report: stable finding codes,
severities and locations such as 'Paragraph 12', 'Slide 3' or 'Data!B3'.
--strict turns structural warnings into errors and requires a document title; heuristic findings stay warnings.
--format json (default) prints the report JSON, sarif prints a SARIF 2.1.0 log for code-scanning tools,
and text prints one line per finding. --output writes the report to a file instead of standard output.
Exit code 0 means no errors were found (and no warnings with --fail-on-warnings); 1 means the policy failed.
The input is opened read-only and never changed.
""";

    internal static async Task<int> RunAsync(string[] args, TextWriter standardOutput, TextWriter standardError,
        CancellationToken cancellationToken = default) {
        AccessibilityArguments parsed;
        try {
            parsed = AccessibilityArguments.Parse(args);
        } catch (ArgumentException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        }
        if (parsed.Help) {
            await standardOutput.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }

        string input = Path.GetFullPath(parsed.Input!);
        if (!File.Exists(input)) {
            await standardError.WriteLineAsync("Input file was not found: " + input).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        }
        if (!OfficeDocument.SupportsAccessibility(input)) {
            await standardError.WriteLineAsync("Accessibility inspection supports Word, Excel and PowerPoint files.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.UnsupportedInput;
        }

        OfficeAccessibilityReport report;
        try {
            report = OfficeDocument.InspectAccessibility(input, parsed.Profile, parsed.Password, cancellationToken);
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Accessibility inspection cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            await standardError.WriteLineAsync("Accessibility inspection failed: " + exception.GetType().Name + ": " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }

        // SARIF keeps the path as given, so a relative input stays a repository-relative artifact URI.
        string text = Render(report, parsed.Format, parsed.Input!);
        if (parsed.Output != null) {
            try {
                string output = Path.GetFullPath(parsed.Output);
                string? directory = Path.GetDirectoryName(output);
                if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
                await File.WriteAllTextAsync(output, text, new UTF8Encoding(false), cancellationToken).ConfigureAwait(false);
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                await standardError.WriteLineAsync("Could not write the report: " + exception.Message).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.OutputFailed;
            }
        } else {
            await standardOutput.WriteLineAsync(text).ConfigureAwait(false);
        }

        bool failed = report.ErrorCount > 0 || parsed.FailOnWarnings && report.WarningCount > 0;
        return (int)(failed ? OfficeImoToolExitCode.ValidationFailed : OfficeImoToolExitCode.Success);
    }

    internal static string Render(OfficeAccessibilityReport report, AccessibilityOutputFormat format, string input) => format switch {
        AccessibilityOutputFormat.Sarif => report.ToSarif(input),
        AccessibilityOutputFormat.Text => RenderText(report),
        _ => report.ToJson()
    };

    private static string RenderText(OfficeAccessibilityReport report) {
        var text = new StringBuilder();
        text.Append(report.Format).Append(" | ").Append(report.Profile).Append(" | ")
            .Append(report.ErrorCount.ToString(CultureInfo.InvariantCulture)).Append(" error(s), ")
            .Append(report.WarningCount.ToString(CultureInfo.InvariantCulture)).Append(" warning(s)");
        foreach (OfficeAccessibilityFinding finding in report.Findings) {
            text.AppendLine().Append("  ").Append(finding.Severity).Append(" | ").Append(finding.Code).Append(" | ")
                .Append(finding.Location);
            if (finding.ObjectName != null) text.Append(" (").Append(finding.ObjectName).Append(')');
            text.Append(" | ").Append(finding.Message);
            if (finding.IsHeuristic) text.Append(" [heuristic]");
        }
        return text.ToString();
    }
}

internal enum AccessibilityOutputFormat {
    Json,
    Sarif,
    Text
}

internal sealed class AccessibilityArguments {
    internal bool Help { get; private set; }
    internal string? Input { get; private set; }
    internal string? Output { get; private set; }
    internal string? Password { get; private set; }
    internal bool FailOnWarnings { get; private set; }
    internal OfficeAccessibilityPolicyProfile Profile { get; private set; } = OfficeAccessibilityPolicyProfile.Default;
    internal AccessibilityOutputFormat Format { get; private set; } = AccessibilityOutputFormat.Json;

    internal static AccessibilityArguments Parse(string[] args) {
        var parsed = new AccessibilityArguments();
        if (args.Length == 0 || args[0] is "help" or "--help" or "-h") {
            parsed.Help = true;
            return parsed;
        }
        for (int index = 0; index < args.Length; index++) {
            string token = args[index];
            switch (token) {
                case "--format":
                    parsed.Format = ParseFormat(Read(args, ref index, token));
                    break;
                case "--strict":
                    parsed.Profile = OfficeAccessibilityPolicyProfile.Strict;
                    break;
                case "--fail-on-warnings":
                    parsed.FailOnWarnings = true;
                    break;
                case "--output":
                    parsed.Output = Read(args, ref index, token);
                    break;
                case "--password":
                    parsed.Password = Read(args, ref index, token);
                    break;
                case "--help" or "-h":
                    parsed.Help = true;
                    return parsed;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal) || parsed.Input != null) throw new ArgumentException("Unexpected argument '" + token + "'.");
                    parsed.Input = token;
                    break;
            }
        }
        if (parsed.Input == null) throw new ArgumentException("An input file is required.");
        return parsed;
    }

    internal static AccessibilityOutputFormat ParseFormat(string value) => value.Trim().ToLowerInvariant() switch {
        "json" => AccessibilityOutputFormat.Json,
        "sarif" => AccessibilityOutputFormat.Sarif,
        "text" => AccessibilityOutputFormat.Text,
        _ => throw new ArgumentException("Unknown format '" + value + "'; use json, sarif or text.")
    };

    private static string Read(string[] args, ref int index, string option) {
        if (index + 1 >= args.Length) throw new ArgumentException(option + " requires a value.");
        return args[++index];
    }
}
