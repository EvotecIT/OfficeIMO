using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OfficeIMO.Processes.Internal;

/// <summary>Launches a Unix process as the leader of a group that survives reparenting of its members.</summary>
internal sealed class OfficeUnixProcessGroup {
    private const string MacPerl = "/usr/bin/perl";
    private const int ExecutePermission = 1;
    private const int Closed = -1;
    private const int SigKill = 9;
    private static readonly Lazy<string> PerlArgumentsPrefix = new(() => "-MPOSIX -e " +
        OfficeProcessArguments.QuoteArgument(
            "defined(POSIX::setsid()) or die qq(setsid failed: $!\\n); exec { $ARGV[0] } @ARGV; die qq(exec failed: $!\\n);") + " -- ");
    private static readonly Lazy<string?> HostExecutablePath = new(GetHostExecutablePath);
    private int _processGroupId;

    private OfficeUnixProcessGroup() { }

    /// <summary>Wraps the executable in an available system session launcher and preserves its argument values.</summary>
    internal static OfficeUnixProcessGroup? TryConfigure(ProcessStartInfo startInfo, bool resolveExecutable = false) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;

        string? launcher = ResolveSetSid();
        string? prefix = launcher != null ? "-- " : null;
        if (launcher == null && RuntimeInformation.IsOSPlatform(OSPlatform.OSX) && File.Exists(MacPerl)) {
            launcher = MacPerl;
            prefix = PerlArgumentsPrefix.Value;
        }
        if (launcher == null) return null;

        string? executable = resolveExecutable ? ResolveExecutable(startInfo.FileName) : startInfo.FileName;
        if (executable == null) return null;
        startInfo.FileName = launcher;
        startInfo.Arguments = prefix + OfficeProcessArguments.QuoteArgument(executable) + " " + startInfo.Arguments;
        return new OfficeUnixProcessGroup();
    }

    /// <summary>Records the PID of the launched group leader unless the lifetime already ended.</summary>
    internal void TrackLeader(int processId) {
        _ = Interlocked.CompareExchange(ref _processGroupId, processId, 0);
    }

    /// <summary>Signals the owned group once, including members whose original parent has exited.</summary>
    internal bool TryTerminate() {
        int processGroupId = Interlocked.Exchange(ref _processGroupId, Closed);
        if (processGroupId <= 0) return false;

        try {
            return Kill(-processGroupId, SigKill) == 0;
        } catch (Exception error) when (error is DllNotFoundException or EntryPointNotFoundException) {
            return false;
        }
    }

    private static string? ResolveExecutable(string fileName) {
        if (Path.IsPathRooted(fileName)) return File.Exists(fileName) ? fileName : null;

        // Match direct Process.Start lookup before the launcher enters the child's working directory.
        string? hostExecutable = HostExecutablePath.Value;
        if (hostExecutable == null) return null;
        string hostCandidate = Path.Combine(Path.GetDirectoryName(hostExecutable)!, fileName);
        if (File.Exists(hostCandidate)) return hostCandidate;
        string currentCandidate = Path.Combine(Directory.GetCurrentDirectory(), fileName);
        if (File.Exists(currentCandidate)) return currentCandidate;

        string? path = Environment.GetEnvironmentVariable("PATH");
        if (path == null) return null;
        foreach (string directory in path.Split(Path.PathSeparator)) {
            if (directory.Length == 0) continue;
            string candidate = Path.Combine(directory, fileName);
            if (!File.Exists(candidate)) continue;
            // Keep direct Process.Start resolution for ambiguous PATH entries.
            return Path.IsPathRooted(candidate) && Access(candidate, ExecutePermission) == 0 ? candidate : null;
        }
        return null;
    }

    private static string? GetHostExecutablePath() {
#if NET8_0_OR_GREATER
        return Environment.ProcessPath;
#else
        return (string?)typeof(Environment).GetProperty("ProcessPath")?.GetValue(null);
#endif
    }

    private static string? ResolveSetSid() {
        if (File.Exists("/usr/bin/setsid")) return "/usr/bin/setsid";
        return File.Exists("/bin/setsid") ? "/bin/setsid" : null;
    }

    [DllImport("libc", EntryPoint = "access")]
    private static extern int Access(string path, int mode);

    [DllImport("libc", EntryPoint = "kill", SetLastError = true)]
    private static extern int Kill(int pid, int signal);
}
