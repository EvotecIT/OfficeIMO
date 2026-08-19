namespace OfficeIMO.OpenDocument;

/// <summary>Configures interoperable password encryption for an OpenDocument package.</summary>
public sealed class OdfEncryptionOptions {
    /// <summary>Password used to encrypt the output package.</summary>
    /// <remarks>The password is encoded as UTF-8, used for this operation, and is not retained.</remarks>
    public string Password { get; set; } = string.Empty;

    /// <summary>PBKDF2-HMAC-SHA1 iteration count applied independently to each encrypted entry.</summary>
    /// <remarks>Values from 10,000 through 10,000,000 are accepted. The default matches current LibreOffice output.</remarks>
    public int IterationCount { get; set; } = 100000;
}
