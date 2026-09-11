using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.ConversionConsistency;

internal static class ArtifactPaths {
    internal static string Resolve(string root, string relative) {
        if (string.IsNullOrWhiteSpace(relative) || Path.IsPathRooted(relative))
            throw new InvalidDataException("Artifact paths must be relative: " + relative);
        string path = Path.GetFullPath(Path.Combine(root, relative));
        string boundary = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!path.StartsWith(boundary, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
            throw new InvalidDataException("Artifact path escapes its root: " + relative);
        return path;
    }

    internal static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    internal static string HashFile(string path) => Hash(File.ReadAllBytes(path));

    internal static async Task<string> RunAsync(string executable, IEnumerable<string> arguments, string? directory = null,
        CancellationToken cancellationToken = default) {
        var start = new ProcessStartInfo(executable) {
            UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true,
            CreateNoWindow = true, WorkingDirectory = directory ?? Environment.CurrentDirectory
        };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Cannot start " + executable);
        Task<string> stdout = process.StandardOutput.ReadToEndAsync(cancellationToken);
        Task<string> stderr = process.StandardError.ReadToEndAsync(cancellationToken);
        try {
            await process.WaitForExitAsync(cancellationToken);
        } catch {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
            throw;
        }
        string output = await stdout;
        string error = await stderr;
        if (process.ExitCode != 0) throw new InvalidOperationException(executable + " failed: " + error + output);
        return (output + error).Trim();
    }

    internal static async Task<(string Commit, string DiffHash, List<SourceFileHash> Untracked)> ProvenanceAsync(string repository) {
        string commit = await RunAsync("git", new[] { "rev-parse", "HEAD" }, repository);
        string diff = await RunAsync("git", new[] { "diff", "HEAD", "--no-ext-diff", "--no-textconv", "--binary" }, repository);
        string untrackedPaths = await RunAsync("git", new[] { "ls-files", "--others", "--exclude-standard", "-z", "--", ".", ":(exclude).artifacts/**" }, repository);
        var untracked = untrackedPaths.Split('\0', StringSplitOptions.RemoveEmptyEntries)
            .OrderBy(path => path, StringComparer.Ordinal)
            .Select(path => new SourceFileHash(path.Replace('\\', '/'), HashFile(Resolve(repository, path)))).ToList();
        return (commit, Hash(Encoding.UTF8.GetBytes(diff)), untracked);
    }
}
