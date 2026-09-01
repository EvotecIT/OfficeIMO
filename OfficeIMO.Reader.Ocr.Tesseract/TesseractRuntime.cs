using System.Runtime.InteropServices;

namespace OfficeIMO.Reader.Ocr.Tesseract;

/// <summary>How a Tesseract executable was located.</summary>
public enum TesseractRuntimeSource {
    /// <summary>The caller supplied an explicit file path.</summary>
    Explicit = 0,
    /// <summary>The <c>OFFICEIMO_TESSERACT_PATH</c> or <c>TESSERACT_PATH</c> environment variable supplied the path.</summary>
    Environment = 1,
    /// <summary>The executable was found on the process search path.</summary>
    Path = 2,
    /// <summary>The executable was found in a known platform installation directory.</summary>
    KnownLocation = 3
}

/// <summary>Resolved local Tesseract runtime information.</summary>
public sealed class TesseractRuntimeInfo {
    internal TesseractRuntimeInfo(string executablePath, string? tessdataDirectory, TesseractRuntimeSource source) {
        ExecutablePath = executablePath;
        TessdataDirectory = tessdataDirectory;
        Source = source;
    }

    /// <summary>Absolute path to the Tesseract executable.</summary>
    public string ExecutablePath { get; }

    /// <summary>Detected trained-data directory, when one could be found without running the executable.</summary>
    public string? TessdataDirectory { get; }

    /// <summary>How the executable was resolved.</summary>
    public TesseractRuntimeSource Source { get; }
}

/// <summary>Dependency-free discovery for an installed Tesseract command-line runtime.</summary>
public static class TesseractRuntime {
    /// <summary>Finds an installed Tesseract executable or throws with platform-specific installation guidance.</summary>
    public static TesseractRuntimeInfo Discover(string? executablePath = null) {
        if (TryDiscover(executablePath, out TesseractRuntimeInfo? runtime)) return runtime!;
        throw new FileNotFoundException(BuildMissingRuntimeMessage(executablePath));
    }

    /// <summary>Attempts to find an installed Tesseract executable without starting a process.</summary>
    public static bool TryDiscover(string? executablePath, out TesseractRuntimeInfo? runtime) {
        runtime = null;
        string? requested = string.IsNullOrWhiteSpace(executablePath) ? null : executablePath!.Trim();
        if (requested != null) {
            if (LooksLikePath(requested)) {
                return TryCreate(requested, TesseractRuntimeSource.Explicit, out runtime);
            }
            if (TryFindOnPath(requested, out string? requestedPath) &&
                TryCreate(requestedPath!, TesseractRuntimeSource.Path, out runtime)) return true;
            return false;
        }

        foreach (string variable in new[] { "OFFICEIMO_TESSERACT_PATH", "TESSERACT_PATH" }) {
            string? value = Environment.GetEnvironmentVariable(variable);
            if (!string.IsNullOrWhiteSpace(value) && TryCreate(value!, TesseractRuntimeSource.Environment, out runtime)) return true;
        }

        string executableName = PlatformExecutableName();
        if (TryFindOnPath(executableName, out string? path) && TryCreate(path!, TesseractRuntimeSource.Path, out runtime)) return true;
        foreach (string candidate in GetKnownLocations()) {
            if (TryCreate(candidate, TesseractRuntimeSource.KnownLocation, out runtime)) return true;
        }
        return false;
    }

    /// <summary>Returns a concise installation command for the current operating system.</summary>
    public static string GetInstallationHint() {
        if (IsWindows()) return "Install Tesseract, for example: winget install --id UB-Mannheim.TesseractOCR --exact";
        if (IsMacOS()) return "Install Tesseract, for example: brew install tesseract";
        return "Install Tesseract with the host package manager, for example: apt-get install tesseract-ocr";
    }

    private static bool TryCreate(string candidate, TesseractRuntimeSource source, out TesseractRuntimeInfo? runtime) {
        runtime = null;
        string expanded = Environment.ExpandEnvironmentVariables(candidate.Trim().Trim('"'));
        string fullPath;
        try {
            fullPath = Path.GetFullPath(expanded);
        } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException || exception is PathTooLongException) {
            return false;
        }
        if (!IsExecutableFile(fullPath)) return false;
        runtime = new TesseractRuntimeInfo(fullPath, FindTessdataDirectory(fullPath), source);
        return true;
    }

    private static bool TryFindOnPath(string executableName, out string? result) {
        return TryFindOnPath(executableName, Environment.GetEnvironmentVariable("PATH"), out result);
    }

    /// <summary>Searches one path value and skips files that the current Unix process cannot execute.</summary>
    internal static bool TryFindOnPath(string executableName, string? path, out string? result) {
        result = null;
        if (string.IsNullOrWhiteSpace(path)) return false;
        string[] extensions = IsWindows()
            ? GetWindowsExecutableExtensions(executableName)
            : new[] { string.Empty };
        foreach (string directory in path!.Split(Path.PathSeparator)) {
            if (string.IsNullOrWhiteSpace(directory)) continue;
            foreach (string extension in extensions) {
                string candidate;
                try {
                    candidate = Path.Combine(directory.Trim().Trim('"'), executableName + extension);
                } catch (ArgumentException) {
                    continue;
                }
                if (IsExecutableFile(candidate)) {
                    result = candidate;
                    return true;
                }
            }
        }
        return false;
    }

    private static string[] GetWindowsExecutableExtensions(string executableName) {
        if (Path.HasExtension(executableName)) return new[] { string.Empty };
        string? pathExt = Environment.GetEnvironmentVariable("PATHEXT");
        string[] values = string.IsNullOrWhiteSpace(pathExt)
            ? new[] { ".exe", ".cmd", ".bat" }
            : pathExt!.Split(';').Where(static extension => !string.IsNullOrWhiteSpace(extension)).ToArray();
        return values.Length == 0 ? new[] { ".exe" } : values;
    }

    private static IEnumerable<string> GetKnownLocations() {
        string executable = PlatformExecutableName();
        if (IsWindows()) {
            foreach (Environment.SpecialFolder folder in new[] { Environment.SpecialFolder.ProgramFiles, Environment.SpecialFolder.ProgramFilesX86, Environment.SpecialFolder.LocalApplicationData }) {
                string root = Environment.GetFolderPath(folder);
                if (string.IsNullOrWhiteSpace(root)) continue;
                yield return folder == Environment.SpecialFolder.LocalApplicationData
                    ? Path.Combine(root, "Programs", "Tesseract-OCR", executable)
                    : Path.Combine(root, "Tesseract-OCR", executable);
            }
            yield break;
        }
        yield return "/usr/bin/tesseract";
        yield return "/usr/local/bin/tesseract";
        if (IsMacOS()) {
            yield return "/opt/homebrew/bin/tesseract";
            yield return "/usr/local/opt/tesseract/bin/tesseract";
        }
    }

    private static string? FindTessdataDirectory(string executablePath) {
        string? directory = Path.GetDirectoryName(executablePath);
        if (!string.IsNullOrWhiteSpace(directory)) {
            string adjacent = Path.Combine(directory!, "tessdata");
            if (Directory.Exists(adjacent)) return adjacent;
        }
        string? prefix = Environment.GetEnvironmentVariable("TESSDATA_PREFIX");
        if (!string.IsNullOrWhiteSpace(prefix)) {
            string direct = Path.GetFullPath(Environment.ExpandEnvironmentVariables(prefix!.Trim().Trim('"')));
            string nested = Path.Combine(direct, "tessdata");
            if (Directory.Exists(nested)) return nested;
            if (Directory.Exists(direct)) return direct;
        }
        foreach (string candidate in new[] { "/usr/share/tesseract-ocr/5/tessdata", "/usr/share/tesseract-ocr/4.00/tessdata", "/usr/share/tessdata" }) {
            if (Directory.Exists(candidate)) return candidate;
        }
        return null;
    }

    private static bool LooksLikePath(string? value) => !string.IsNullOrWhiteSpace(value) &&
        (Path.IsPathRooted(value) || value!.IndexOf(Path.DirectorySeparatorChar) >= 0 || value.IndexOf(Path.AltDirectorySeparatorChar) >= 0);

    private static string PlatformExecutableName() => IsWindows() ? "tesseract.exe" : "tesseract";

    private static string BuildMissingRuntimeMessage(string? requested) {
        string prefix = string.IsNullOrWhiteSpace(requested)
            ? "Tesseract was not found in the OfficeIMO/Tesseract environment variables, process PATH, or known installation locations."
            : "Tesseract executable '" + requested!.Trim() + "' was not found.";
        return prefix + " " + GetInstallationHint() + ". Supply TesseractOcrEngineOptions.ExecutablePath when it is installed elsewhere.";
    }

    private static bool IsWindows() => Environment.OSVersion.Platform == PlatformID.Win32NT;
    private static bool IsMacOS() => Environment.OSVersion.Platform == PlatformID.MacOSX || Directory.Exists("/Applications") && Directory.Exists("/System");

    private static bool IsExecutableFile(string path) =>
        File.Exists(path) && (IsWindows() || Access(path, ExecutePermission) == 0);

    private const int ExecutePermission = 1;

    [DllImport("libc", EntryPoint = "access", SetLastError = true)]
    private static extern int Access(string path, int mode);
}
