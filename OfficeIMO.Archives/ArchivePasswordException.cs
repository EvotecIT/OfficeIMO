namespace OfficeIMO.Archives;

/// <summary>An encrypted archive member could not be opened: no password was supplied, or the supplied password does
/// not open it.</summary>
/// <remarks>Derives from <see cref="IOException"/>, like <see cref="ArchiveCorruptionException"/>. Traditional ZipCrypto
/// carries only a one-byte password check, so a wrong password can also surface after decoding, as a checksum
/// mismatch; that case is reported here too, because it cannot be told apart from a damaged member.</remarks>
public sealed class ArchivePasswordException : IOException {
    /// <summary>Creates the exception for <paramref name="entryName"/>.</summary>
    /// <param name="entryName">The member name as stored in the archive.</param>
    /// <param name="passwordSupplied">Whether a password was supplied.</param>
    /// <param name="message">The error message.</param>
    public ArchivePasswordException(string entryName, bool passwordSupplied, string message)
        : this(entryName, passwordSupplied, message, null) {
    }

    /// <summary>Creates the exception with the error that caused it.</summary>
    /// <param name="entryName">The member name as stored in the archive.</param>
    /// <param name="passwordSupplied">Whether a password was supplied.</param>
    /// <param name="message">The error message.</param>
    /// <param name="innerException">The error that caused this one, or null.</param>
    public ArchivePasswordException(string entryName, bool passwordSupplied, string message, Exception? innerException)
        : base(message, innerException) {
        EntryName = entryName ?? throw new ArgumentNullException(nameof(entryName));
        PasswordSupplied = passwordSupplied;
    }

    /// <summary>The member name as stored in the archive.</summary>
    public string EntryName { get; }

    /// <summary>True when a password was supplied but does not open the member; false when none was supplied.</summary>
    public bool PasswordSupplied { get; }
}
