using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace OfficeIMO.Markup.PowerPoint;

internal static class OfficeMarkupMermaidRenderer {
    private const string MermaidEnvironmentVariable = "OFFICEIMO_MARKUP_MERMAID_CLI";

    public static bool TryRenderPng(
        OfficeMarkupDiagramBlock diagram,
        OfficeMarkupPowerPointExportOptions options,
        out string outputPath) {
        outputPath = string.Empty;
        if (diagram == null || options == null || !options.RenderMermaidDiagrams || !diagram.RenderAsImage) {
            return false;
        }

        if (!IsMermaid(diagram.Language) || string.IsNullOrWhiteSpace(diagram.Content)) {
            return false;
        }

        var renderer = ResolveRendererPath(options);
        if (string.IsNullOrWhiteSpace(renderer)) {
            return false;
        }

        var tempDirectory = ResolveTemporaryDirectory(options);
        Directory.CreateDirectory(tempDirectory);

        var stem = "officeimo-mermaid-" + Guid.NewGuid().ToString("N");
        var inputPath = Path.Combine(tempDirectory, stem + ".mmd");
        var configPath = Path.Combine(tempDirectory, stem + ".json");
        outputPath = Path.Combine(tempDirectory, stem + ".png");
        File.WriteAllText(inputPath, diagram.Content, Encoding.UTF8);
        File.WriteAllText(configPath, MermaidConfigJson, Encoding.UTF8);

        try {
            var themedArguments = "-i " + Quote(inputPath)
                + " -o " + Quote(outputPath)
                + " -b transparent"
                + " -c " + Quote(configPath)
                + " -s 2";
            if (TryRunRenderer(renderer!, themedArguments, outputPath, options)) {
                return true;
            }

            TryDelete(outputPath);
            var fallbackArguments = "-i " + Quote(inputPath) + " -o " + Quote(outputPath) + " -b transparent";
            return TryRunRenderer(renderer!, fallbackArguments, outputPath, options);
        } catch (Exception) when (!Debugger.IsAttached) {
            return false;
        } finally {
            TryDelete(inputPath);
            TryDelete(configPath);
        }
    }

    private static bool TryRunRenderer(
        string renderer,
        string arguments,
        string outputPath,
        OfficeMarkupPowerPointExportOptions options) {
        var processStartInfo = CreateRendererStartInfo(renderer, arguments);

        using var process = Process.Start(processStartInfo);
        if (process == null) {
            return false;
        }

        var standardOutputTask = process.StandardOutput.ReadToEndAsync();
        var standardErrorTask = process.StandardError.ReadToEndAsync();
        var timeout = Math.Max(1000, options.MermaidRenderTimeoutMilliseconds);
        if (!process.WaitForExit(timeout)) {
            try {
                process.Kill();
            } catch (InvalidOperationException ex) {
                Debug.WriteLine($"OfficeIMO.Markup.Mermaid renderer already exited while being terminated: {ex.Message}");
            }

            return false;
        }

        process.WaitForExit();
        Task.WaitAll(new Task[] { standardOutputTask, standardErrorTask }, timeout);
        return process.ExitCode == 0 && File.Exists(outputPath);
    }

    private const string MermaidConfigJson = """
{
  "theme": "base",
  "themeVariables": {
    "fontFamily": "Aptos, Arial, sans-serif",
    "fontSize": "18px",
    "primaryColor": "#EEF4FF",
    "primaryBorderColor": "#A5B4FC",
    "primaryTextColor": "#172033",
    "lineColor": "#64748B",
    "tertiaryColor": "#F8FAFC"
  },
  "flowchart": {
    "htmlLabels": true,
    "curve": "basis",
    "nodeSpacing": 42,
    "rankSpacing": 56
  }
}
""";

    private static string? ResolveRendererPath(OfficeMarkupPowerPointExportOptions options) {
        if (!string.IsNullOrWhiteSpace(options.MermaidRendererPath)) {
            return options.MermaidRendererPath;
        }

        var environmentPath = Environment.GetEnvironmentVariable(MermaidEnvironmentVariable);
        if (!string.IsNullOrWhiteSpace(environmentPath)) {
            return environmentPath;
        }

        return FindExecutableInPath("mmdc")
            ?? FindBundledRenderer();
    }

    private static string? FindBundledRenderer() {
        var installRoot = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markup.Mermaid", "node_modules", ".bin");
        return FindExecutableInDirectory(installRoot, "mmdc");
    }

    private static ProcessStartInfo CreateRendererStartInfo(string renderer, string arguments) {
        var fileName = renderer;
        var processArguments = arguments;
        if (IsWindowsBatchFile(renderer)) {
            fileName = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
            processArguments = "/d /s /c \"" + Quote(renderer) + " " + arguments + "\"";
        }

        return new ProcessStartInfo {
            FileName = fileName,
            Arguments = processArguments,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardError = true,
            RedirectStandardOutput = true
        };
    }

    private static bool IsWindowsBatchFile(string path) =>
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
        && (path.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase)
            || path.EndsWith(".bat", StringComparison.OrdinalIgnoreCase));

    private static string ResolveTemporaryDirectory(OfficeMarkupPowerPointExportOptions options) {
        if (!string.IsNullOrWhiteSpace(options.TemporaryDirectory)) {
            return options.TemporaryDirectory!;
        }

        return Path.Combine(Path.GetTempPath(), "OfficeIMO.Markup");
    }

    private static string? FindExecutableInPath(string command) {
        var paths = (Environment.GetEnvironmentVariable("PATH") ?? string.Empty)
            .Split(Path.PathSeparator)
            .Where(path => !string.IsNullOrWhiteSpace(path));

        foreach (var path in paths) {
            var candidate = FindExecutableInDirectory(path.Trim(), command);
            if (candidate != null) {
                return candidate;
            }
        }

        return null;
    }

    private static string? FindExecutableInDirectory(string directory, string command) {
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory)) {
            return null;
        }

        var safeCommand = Path.GetFileName(command);
        if (string.IsNullOrWhiteSpace(safeCommand)) {
            return null;
        }

        var extensions = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? new[] { ".cmd", ".bat", ".exe", ".com", ".ps1", string.Empty }
            : new[] { string.Empty };

        foreach (var extension in extensions) {
            var candidate = Path.Combine(directory, safeCommand + extension);
            if (File.Exists(candidate)) {
                return candidate;
            }
        }

        return null;
    }

    private static bool IsMermaid(string language) =>
        string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase);

    private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";

    private static void TryDelete(string path) {
        try {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        } catch (IOException ex) {
            Trace.TraceWarning($"OfficeIMO.Markup.Mermaid could not delete temporary file '{path}': {ex.Message}");
        } catch (UnauthorizedAccessException ex) {
            Trace.TraceWarning($"OfficeIMO.Markup.Mermaid could not delete temporary file '{path}': {ex.Message}");
        }
    }
}
