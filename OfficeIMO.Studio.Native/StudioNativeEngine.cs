using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Native;

/// <summary>
/// Process-wide engine settings. The host names its own data folder (inside the app container on Apple platforms)
/// before opening documents, so recovery copies never land in a location the host does not own.
/// </summary>
internal static class StudioNativeEngine {
    private static readonly object Sync = new();
    private static string? _dataDirectory;

    internal static void Initialize(string dataDirectory) {
        if (string.IsNullOrWhiteSpace(dataDirectory) || !Path.IsPathFullyQualified(dataDirectory))
            throw new ArgumentException("Give the engine an absolute folder the app owns.", nameof(dataDirectory));
        Directory.CreateDirectory(dataDirectory);
        lock (Sync) _dataDirectory = Path.GetFullPath(dataDirectory);
    }

    internal static PdfWorkspaceRecoveryStore CreateRecoveryStore() {
        string root;
        lock (Sync) root = _dataDirectory ?? throw new StudioNativeNotInitializedException();
        return new PdfWorkspaceRecoveryStore(Path.Combine(root, "Recovery"));
    }
}

internal sealed class StudioNativeNotInitializedException()
    : InvalidOperationException("Call officeimo_initialize with the app's data folder before opening documents.");
