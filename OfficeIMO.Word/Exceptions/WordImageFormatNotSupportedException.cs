namespace OfficeIMO.Word;

/// <summary>
/// Exception thrown when an unsupported image format is encountered.
/// </summary>
public class WordImageFormatNotSupportedException : WordException {
    /// <summary>
    /// Initializes a new instance of the <see cref="WordImageFormatNotSupportedException"/> class.
    /// </summary>
    /// <param name="message">Exception message.</param>
    public WordImageFormatNotSupportedException(string message) : base(message) {

    }
}
