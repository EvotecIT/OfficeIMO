using System.ComponentModel;
using System.IO;
using System.Threading;

namespace OfficeIMO.Provenance.C2pa;

/// <summary>Readiness of the explicitly configured tool. Availability does not establish credential or signer trust.</summary>
public sealed class C2paToolAvailability {
    internal C2paToolAvailability(bool available, string executablePath, string? version, string diagnostic) {
        Available = available;
        ExecutablePath = executablePath;
        Version = version;
        Diagnostic = diagnostic;
    }
    /// <summary>Whether the configured executable returned a recognized successful version response through the bounded process runner.</summary>
    public bool Available { get; }
    /// <summary>Gets the host-supplied executable selection.</summary>
    public string ExecutablePath { get; }
    /// <summary>Gets the observed tool version, when recognized.</summary>
    public string? Version { get; }
    /// <summary>Gets an actionable readiness diagnostic, separate from asset verification.</summary>
    public string Diagnostic { get; }
}

public sealed partial class C2paToolProvenanceVerifier {
    private static readonly string[] AvailabilityArguments = { "--version" };
    /// <summary>Probes the configured executable and process-containment prerequisites without reading an asset or fetching credentials.</summary>
    /// <remarks>The tool is executed with --version. Only configure an executable you trust. This does not validate trust lists or signing credentials.</remarks>
    public C2paToolAvailability CheckAvailability(TimeSpan? timeout = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        TimeSpan limit = timeout ?? TimeSpan.FromSeconds(10);
        if (limit <= TimeSpan.Zero || limit > TimeSpan.FromMinutes(5)) throw new ArgumentOutOfRangeException(nameof(timeout));
        try {
            C2paToolProcessResult process = _runner.Run(new C2paToolProcessRequest(
                ExecutablePath, AvailabilityArguments, Directory.GetCurrentDirectory(), limit, 4096), cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            const string prefix = "c2patool ";
            string output = process.StandardOutput.Trim();
            if (process.ExitCode != 0) return new C2paToolAvailability(false, ExecutablePath, null,
                "c2patool --version exited with code " + process.ExitCode + ". " + process.StandardError.Trim());
            if (!output.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ||
                !System.Version.TryParse(output.Substring(prefix.Length).Trim(), out System.Version? version))
                return new C2paToolAvailability(false, ExecutablePath, null, "The executable did not return a recognized c2patool version.");
            return new C2paToolAvailability(true, ExecutablePath, version.ToString(),
                "The tool and process containment are available. Asset verification and trust configuration are separate checks.");
        } catch (Exception exception) when (exception is Win32Exception or IOException or InvalidDataException or TimeoutException or InvalidOperationException) {
            return new C2paToolAvailability(false, ExecutablePath, null, exception.Message);
        }
    }
}
