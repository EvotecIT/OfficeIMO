namespace OfficeIMO.Email.Store;

/// <summary>Raised when a PST password is required or does not match the store checksum.</summary>
public sealed class EmailStorePasswordException : UnauthorizedAccessException {
    internal EmailStorePasswordException(bool passwordWasProvided)
        : base(passwordWasProvided
            ? "The supplied PST password does not match the store password checksum."
            : "The PST is password protected; provide a password in EmailStoreReaderOptions.") {
        PasswordWasProvided = passwordWasProvided;
    }

    /// <summary>True when validation failed for a supplied password; false when no password was supplied.</summary>
    public bool PasswordWasProvided { get; }
}
