using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

/// <summary>A source retained by a host that must remain separate from workflow outputs even when it is not executed.</summary>
public sealed class OfficeWorkflowProtectedSource {
    /// <summary>Creates source protection with optional provider access for checking the local file identity.</summary>
    public OfficeWorkflowProtectedSource(string location, OfficeWorkflowStreamInput? inputStream = null) {
        Location = OfficeStorageIdentity.Normalize(location);
        InputStream = inputStream;
    }
    /// <summary>Original local path or provider location to protect.</summary>
    public string Location { get; }
    /// <summary>Reopenable provider access retained by the host for this source.</summary>
    public OfficeWorkflowStreamInput? InputStream { get; }
}
