namespace OfficeIMO.Word;

/// <summary>
/// Base class for all OfficeIMO specific exceptions.
/// </summary>
public abstract class OfficeIMOException : Exception {
    /// <summary>
    /// Initializes a new instance of the <see cref="OfficeIMOException"/> class.
    /// </summary>
    /// <param name="message">Exception message.</param>
    protected OfficeIMOException(string message) : base(message) { }
}
