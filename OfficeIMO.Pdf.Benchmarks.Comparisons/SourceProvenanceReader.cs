using System.Diagnostics;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed record GitSourceState(string Commit, bool IsClean);

internal static class SourceProvenanceReader {
    internal static async Task<GitSourceState?> ReadGitStateAsync(
        string path,
        CancellationToken cancellationToken = default) {
        string? root = await RunGitAsync(path, new[] { "rev-parse", "--show-toplevel" }, cancellationToken).ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(root)) return null;
        string? commit = await RunGitAsync(root, new[] { "rev-parse", "HEAD" }, cancellationToken).ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(commit)) return null;
        string? status = await RunGitAsync(
            root,
            new[] { "status", "--porcelain", "--untracked-files=normal" },
            cancellationToken).ConfigureAwait(false);
        return status == null ? null : new GitSourceState(commit.Trim(), string.IsNullOrWhiteSpace(status));
    }

    private static async Task<string?> RunGitAsync(
        string workingDirectory,
        IReadOnlyList<string> arguments,
        CancellationToken cancellationToken) {
        var startInfo = new ProcessStartInfo("git") {
            WorkingDirectory = workingDirectory,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in arguments) startInfo.ArgumentList.Add(argument);
        try {
            using Process process = Process.Start(startInfo)
                ?? throw new InvalidOperationException("Could not start Git for source provenance.");
            string output = await process.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
            _ = await process.StandardError.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
            await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
            return process.ExitCode == 0 ? output.Trim() : null;
        } catch (System.ComponentModel.Win32Exception) {
            return null;
        }
    }
}
