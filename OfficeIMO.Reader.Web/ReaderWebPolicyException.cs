namespace OfficeIMO.Reader.Web;

/// <summary>Thrown when a requested or reported response URI violates Reader Web policy.</summary>
public sealed class ReaderWebPolicyException : InvalidOperationException {
    internal ReaderWebPolicyException(Uri targetUri, string message)
        : base(message) {
        TargetUri = targetUri;
    }

    /// <summary>Gets the URI rejected by policy.</summary>
    public Uri TargetUri { get; }
}
