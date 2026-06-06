using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Tests.Pdf;

internal sealed class PdfExternalValidator {
    private const string RequireEnv = "OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS";
    private readonly string[] _arguments;

    private PdfExternalValidator(string name, string? executablePath, string[] arguments, bool autoDetected) {
        Name = name;
        ExecutablePath = executablePath;
        _arguments = arguments;
        AutoDetected = autoDetected;
    }

    internal string Name { get; }

    internal string? ExecutablePath { get; }

    internal bool AutoDetected { get; }

    internal bool IsAvailable => !string.IsNullOrWhiteSpace(ExecutablePath);

    internal string GetNotConfiguredText() =>
        Name + " was not configured, so external validation was not run." + Environment.NewLine +
        "Set OFFICEIMO_VERAPDF, OFFICEIMO_VERAPDF_PATH, OFFICEIMO_PDFUA_VALIDATOR, OFFICEIMO_PDFUA_VALIDATOR_PATH, OFFICEIMO_MUSTANG, or OFFICEIMO_MUSTANG_PATH as appropriate, or add the tool to PATH." + Environment.NewLine;

    internal static PdfExternalValidator VeraPdf() {
        string? explicitPath = FirstNonEmpty(
            Environment.GetEnvironmentVariable("OFFICEIMO_VERAPDF"),
            Environment.GetEnvironmentVariable("OFFICEIMO_VERAPDF_PATH"));
        string? path = explicitPath ?? FindOnPath("verapdf", "verapdf.bat", "verapdf.exe");
        string[] args = GetConfiguredArgs("OFFICEIMO_VERAPDF_ARGS", "-f", "3b", "{pdf}");
        return new PdfExternalValidator("veraPDF", path, args, explicitPath == null && path != null);
    }

    internal static PdfExternalValidator PdfUa() {
        string? explicitPath = FirstNonEmpty(
            Environment.GetEnvironmentVariable("OFFICEIMO_PDFUA_VALIDATOR"),
            Environment.GetEnvironmentVariable("OFFICEIMO_PDFUA_VALIDATOR_PATH"));
        string? path = explicitPath ?? FindOnPath("pdfua-validator", "pdfua-validator.bat", "pdfua-validator.exe");
        string[] args = GetConfiguredArgs("OFFICEIMO_PDFUA_VALIDATOR_ARGS", "{pdf}");
        if (path != null && string.Equals(Path.GetExtension(path), ".jar", StringComparison.OrdinalIgnoreCase)) {
            args = new[] { "-jar", path }.Concat(args).ToArray();
            path = FindOnPath("java", "java.exe");
        }

        return new PdfExternalValidator("PDF/UA validator", path, args, explicitPath == null && path != null);
    }

    internal static PdfExternalValidator Mustang() {
        string? explicitPath = FirstNonEmpty(
            Environment.GetEnvironmentVariable("OFFICEIMO_MUSTANG"),
            Environment.GetEnvironmentVariable("OFFICEIMO_MUSTANG_PATH"));

        string? path = explicitPath ?? FindOnPath("mustangproject", "mustangproject.bat", "mustangproject.exe", "mustang", "mustang.bat", "mustang.exe");
        string[] args = GetConfiguredArgs("OFFICEIMO_MUSTANG_ARGS", "--action", "validate", "--source", "{pdf}");
        if (path != null && string.Equals(Path.GetExtension(path), ".jar", StringComparison.OrdinalIgnoreCase)) {
            args = new[] { "-jar", path }.Concat(args).ToArray();
            path = FindOnPath("java", "java.exe");
        }

        return new PdfExternalValidator("Mustang", path, args, explicitPath == null && path != null);
    }

    internal static void SkipUnlessRequired(PdfExternalValidator validator) {
        if (IsRequired()) {
            throw new InvalidOperationException(
                validator.Name + " compliance validation was required, but the validator was not found. Set OFFICEIMO_VERAPDF, OFFICEIMO_VERAPDF_PATH, OFFICEIMO_PDFUA_VALIDATOR, OFFICEIMO_PDFUA_VALIDATOR_PATH, OFFICEIMO_MUSTANG, or OFFICEIMO_MUSTANG_PATH as appropriate, or add the tool to PATH.");
        }
    }

    internal PdfExternalProcessResult Run(byte[] pdfBytes, string fileName) {
        if (!IsAvailable) {
            throw new InvalidOperationException(Name + " validator is not configured.");
        }

        string workDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.PdfCompliance", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);
        string pdfPath = Path.Combine(workDir, fileName);
        try {
            File.WriteAllBytes(pdfPath, pdfBytes);
            string arguments = string.Join(" ", _arguments.Select(argument => QuoteArgument(argument.Replace("{pdf}", pdfPath))));
            var startInfo = new ProcessStartInfo {
                FileName = ExecutablePath!,
                Arguments = arguments,
                WorkingDirectory = workDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();
            using var process = new Process {
                StartInfo = startInfo
            };
            process.OutputDataReceived += (_, e) => {
                if (e.Data != null) {
                    outputBuilder.AppendLine(e.Data);
                }
            };
            process.ErrorDataReceived += (_, e) => {
                if (e.Data != null) {
                    errorBuilder.AppendLine(e.Data);
                }
            };

            if (!process.Start()) {
                throw new InvalidOperationException("Failed to start " + Name + " validator.");
            }

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            if (!process.WaitForExit(60000)) {
                try {
                    process.Kill();
                } catch (InvalidOperationException) {
                }

                throw new TimeoutException(Name + " validation did not finish within 60 seconds.");
            }

            process.WaitForExit();
            string output = outputBuilder.ToString();
            string error = errorBuilder.ToString();
            return new PdfExternalProcessResult(Name, ExecutablePath!, arguments, process.ExitCode, output, error, AutoDetected);
        } finally {
            TryDeleteDirectory(workDir);
        }
    }

    private static bool IsRequired() =>
        string.Equals(Environment.GetEnvironmentVariable(RequireEnv), "1", StringComparison.Ordinal);

    private static string[] GetConfiguredArgs(string envName, params string[] defaultArgs) {
        string? raw = Environment.GetEnvironmentVariable(envName);
        return string.IsNullOrWhiteSpace(raw)
            ? defaultArgs
            : SplitCommandLine(raw!);
    }

    private static string? FirstNonEmpty(params string?[] values) {
        foreach (string? value in values) {
            if (!string.IsNullOrWhiteSpace(value)) {
                return value;
            }
        }

        return null;
    }

    private static string? FindOnPath(params string[] names) {
        string? path = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(path)) {
            return null;
        }

        foreach (string directory in path.Split(Path.PathSeparator)) {
            if (string.IsNullOrWhiteSpace(directory)) {
                continue;
            }

            foreach (string name in names) {
                string candidate = Path.Combine(directory, name);
                if (File.Exists(candidate)) {
                    return candidate;
                }
            }
        }

        return null;
    }

    private static string[] SplitCommandLine(string value) {
        var args = new List<string>();
        var current = new StringBuilder();
        bool inQuotes = false;

        foreach (char c in value) {
            if (c == '"') {
                inQuotes = !inQuotes;
                continue;
            }

            if (char.IsWhiteSpace(c) && !inQuotes) {
                if (current.Length > 0) {
                    args.Add(current.ToString());
                    current.Clear();
                }

                continue;
            }

            current.Append(c);
        }

        if (current.Length > 0) {
            args.Add(current.ToString());
        }

        return args.ToArray();
    }

    private static string QuoteArgument(string argument) {
        if (argument.Length == 0) {
            return "\"\"";
        }

        if (argument.IndexOfAny(new[] { ' ', '\t', '"' }) < 0) {
            return argument;
        }

        var builder = new StringBuilder(argument.Length + 2);
        builder.Append('"');
        int backslashCount = 0;
        foreach (char c in argument) {
            if (c == '\\') {
                backslashCount++;
                continue;
            }

            if (c == '"') {
                builder.Append('\\', (backslashCount * 2) + 1);
                builder.Append('"');
                backslashCount = 0;
                continue;
            }

            if (backslashCount > 0) {
                builder.Append('\\', backslashCount);
                backslashCount = 0;
            }

            builder.Append(c);
        }

        if (backslashCount > 0) {
            builder.Append('\\', backslashCount * 2);
        }

        builder.Append('"');
        return builder.ToString();
    }

    private static void TryDeleteDirectory(string path) {
        try {
            if (Directory.Exists(path)) {
                Directory.Delete(path, recursive: true);
            }
        } catch (IOException) {
        } catch (UnauthorizedAccessException) {
        }
    }
}

internal sealed class PdfExternalProcessResult {
    internal PdfExternalProcessResult(string validatorName, string executablePath, string arguments, int exitCode, string output, string error, bool autoDetected) {
        ValidatorName = validatorName;
        ExecutablePath = executablePath;
        Arguments = arguments;
        ExitCode = exitCode;
        Output = output;
        Error = error;
        AutoDetected = autoDetected;
    }

    internal string ValidatorName { get; }

    internal string ExecutablePath { get; }

    internal string Arguments { get; }

    internal int ExitCode { get; }

    internal string Output { get; }

    internal string Error { get; }

    internal bool AutoDetected { get; }

    internal string GetDiagnosticText() {
        var sb = new StringBuilder();
        sb.Append(ValidatorName)
            .Append(" exited with code ")
            .Append(ExitCode.ToString(CultureInfo.InvariantCulture))
            .Append(" using ")
            .Append(ExecutablePath)
            .Append(' ')
            .Append(Arguments)
            .AppendLine();
        if (!string.IsNullOrWhiteSpace(Output)) {
            sb.AppendLine(Output.Trim());
        }

        if (!string.IsNullOrWhiteSpace(Error)) {
            sb.AppendLine(Error.Trim());
        }

        return sb.ToString();
    }
}
