namespace OfficeIMO.Word;

/// <summary>
/// Base class for all OfficeIMO specific exceptions.
/// </summary>
public abstract class WordException : Exception {
    /// <summary>
    /// Initializes a new instance of the <see cref="WordException"/> class.
    /// </summary>
    /// <param name="message">Exception message.</param>
    protected WordException(string message) : base(message) { }
}
