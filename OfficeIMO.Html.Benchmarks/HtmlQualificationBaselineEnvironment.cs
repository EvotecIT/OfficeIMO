using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeIMO.Html.Benchmarks;

internal static partial class HtmlQualificationBaselineRunner {
    private static HtmlQualificationEnvironmentEvidence BuildEnvironmentEvidence() {
        using Process process = Process.GetCurrentProcess();
        return new HtmlQualificationEnvironmentEvidence(
            RuntimeInformation.FrameworkDescription,
            RuntimeInformation.OSDescription,
            RuntimeInformation.ProcessArchitecture.ToString(),
            ResolveProcessorName(),
            ResolvePhysicalCoreCount(),
            Environment.ProcessorCount,
            TryGet(() => process.PriorityClass.ToString()),
            ResolveProcessAffinity(process),
            ResolvePowerPlan());
    }

    private static IReadOnlyList<HtmlQualificationProviderEvidence> BuildProviderEvidence() =>
        AppDomain.CurrentDomain.GetAssemblies()
            .Where(assembly => {
                string name = assembly.GetName().Name ?? string.Empty;
                return name.StartsWith("OfficeIMO.Html", StringComparison.Ordinal)
                    || name is "OfficeIMO.Core" or "OfficeIMO.Pdf" or "AngleSharp" or "AngleSharp.Css";
            })
            .OrderBy(assembly => assembly.GetName().Name, StringComparer.Ordinal)
            .Select(assembly => new HtmlQualificationProviderEvidence(
                assembly.GetName().Name ?? "unknown",
                assembly.GetName().Version?.ToString() ?? "unknown",
                assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "unknown"))
            .ToArray();

    private static string ResolveSourceRoot(string? explicitRoot) {
        if (!string.IsNullOrWhiteSpace(explicitRoot)) return ValidateSourceRoot(explicitRoot!);
        string? current = FindSourceRoot(Directory.GetCurrentDirectory());
        if (current != null) return current;
        current = FindSourceRoot(AppContext.BaseDirectory);
        if (current != null) return current;
        throw new InvalidOperationException("OfficeIMO source root was not found. Run from the repository or pass --source-root.");
    }

    private static string? FindSourceRoot(string startPath) {
        var directory = new DirectoryInfo(Path.GetFullPath(startPath));
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Html.Benchmarks", "OfficeIMO.Html.Benchmarks.csproj")) &&
                (Directory.Exists(Path.Combine(directory.FullName, ".git")) || File.Exists(Path.Combine(directory.FullName, ".git")))) {
                return directory.FullName;
            }
            directory = directory.Parent;
        }
        return null;
    }

    private static string ValidateSourceRoot(string root) => FindSourceRoot(root) is string resolved &&
        string.Equals(Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            resolved.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), PathComparison())
        ? resolved
        : throw new InvalidOperationException("The supplied --source-root is not an OfficeIMO repository root: " + root + ".");

    private static string ResolveCommit(string sourceRoot) {
        string value = RunGit(sourceRoot, "rev-parse", "HEAD").Trim();
        if (value.Length == 0) throw new InvalidOperationException("Git provenance could not be read from " + sourceRoot + ".");
        return value;
    }

    private static bool ResolveTrackedSourceTreeDirty(string sourceRoot) =>
        !string.IsNullOrWhiteSpace(RunGit(sourceRoot, "status", "--porcelain", "--untracked-files=no"));

    private static string[] ResolveUntrackedSourcePaths(string sourceRoot) => RunGit(sourceRoot, "ls-files", "--others", "--exclude-standard")
        .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

    private static IReadOnlyList<string> ResolveUntrackedSourceRoots(IEnumerable<string> paths) => paths
        .Select(value => value.Replace('\\', '/').Split('/'))
        .Select(parts => string.Join("/", parts.Take(Math.Min(parts.Length, 2))))
        .Distinct(StringComparer.Ordinal)
        .OrderBy(value => value, StringComparer.Ordinal)
        .ToArray();

    private static string RunGit(string sourceRoot, params string[] arguments) {
        try {
            var info = new ProcessStartInfo("git") {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = sourceRoot
            };
            foreach (string argument in arguments) info.ArgumentList.Add(argument);
            using Process? process = Process.Start(info);
            if (process == null) return string.Empty;
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            return process.ExitCode == 0 ? output : string.Empty;
        } catch {
            return string.Empty;
        }
    }

    private static StringComparison PathComparison() => RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
        ? StringComparison.OrdinalIgnoreCase
        : StringComparison.Ordinal;

    private static string ResolveProcessorName() {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            try {
                using Microsoft.Win32.RegistryKey? key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                    @"HARDWARE\DESCRIPTION\System\CentralProcessor\0");
                if (key?.GetValue("ProcessorNameString") is string name && !string.IsNullOrWhiteSpace(name)) return name.Trim();
            } catch {
                // Fall through to the environment value when registry metadata is unavailable.
            }
        }
        string? identifier = Environment.GetEnvironmentVariable("PROCESSOR_IDENTIFIER");
        if (!string.IsNullOrWhiteSpace(identifier)) return identifier.Trim();
        try {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) && File.Exists("/proc/cpuinfo")) {
                string? model = File.ReadLines("/proc/cpuinfo").FirstOrDefault(line => line.StartsWith("model name", StringComparison.OrdinalIgnoreCase));
                if (model != null && model.IndexOf(':') >= 0) return model.Substring(model.IndexOf(':') + 1).Trim();
            }
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return RunProcess("sysctl", "-n", "machdep.cpu.brand_string");
        } catch {
            // Fall through to the architecture when host topology metadata is unavailable.
        }
        return RuntimeInformation.ProcessArchitecture.ToString();
    }

    private static string ResolvePhysicalCoreCount() {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) && File.Exists("/proc/cpuinfo")) {
            try {
                var cores = new HashSet<string>(StringComparer.Ordinal);
                string physical = "0";
                string core = string.Empty;
                foreach (string line in File.ReadLines("/proc/cpuinfo").Concat(new[] { string.Empty })) {
                    if (line.StartsWith("physical id", StringComparison.OrdinalIgnoreCase)) physical = ValueAfterColon(line);
                    if (line.StartsWith("core id", StringComparison.OrdinalIgnoreCase)) core = ValueAfterColon(line);
                    if (line.Length == 0 && core.Length > 0) {
                        cores.Add(physical + ":" + core);
                        core = string.Empty;
                    }
                }
                if (cores.Count > 0) return cores.Count.ToString(CultureInfo.InvariantCulture);
            } catch {
                // Fall through to an explicit unavailable value.
            }
        }
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return RunProcess("sysctl", "-n", "hw.physicalcpu");
        return "unavailable";
    }

    private static string ResolveProcessAffinity(Process process) {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && !RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            return "not-supported";
#pragma warning disable CA1416 // The runtime guard above limits this call to the two supported platforms.
        return TryGet(() => "0x" + process.ProcessorAffinity.ToInt64().ToString("X", CultureInfo.InvariantCulture));
#pragma warning restore CA1416
    }

    private static string ValueAfterColon(string line) {
        int separator = line.IndexOf(':');
        return separator >= 0 ? line.Substring(separator + 1).Trim() : string.Empty;
    }

    private static string ResolvePowerPlan() => RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
        ? RunProcess("powercfg", "/getactivescheme")
        : "not-applicable";

    private static string RunProcess(string fileName, params string[] arguments) {
        try {
            var info = new ProcessStartInfo(fileName) {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            foreach (string argument in arguments) info.ArgumentList.Add(argument);
            using Process? process = Process.Start(info);
            if (process == null) return "unavailable";
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 && output.Length > 0 ? output : "unavailable";
        } catch {
            return "unavailable";
        }
    }

    private static string TryGet(Func<string> value) {
        try { return value(); }
        catch { return "unavailable"; }
    }
}
