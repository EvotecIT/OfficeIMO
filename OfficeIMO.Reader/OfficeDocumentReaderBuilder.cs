using System;

namespace OfficeIMO.Reader;

/// <summary>
/// Builds an <see cref="OfficeDocumentReader"/> with an isolated handler configuration.
/// </summary>
/// <remarks>
/// A builder may be reused or changed after <see cref="Build"/>. Each built reader retains its own
/// immutable snapshot and is unaffected by later builder or static <see cref="DocumentReader"/> registrations.
/// </remarks>
public sealed class OfficeDocumentReaderBuilder {
    private readonly ReaderHandlerRegistry _handlers = new ReaderHandlerRegistry(DocumentReader.BuiltInExtensions);

    /// <summary>
    /// Adds a handler to this reader configuration.
    /// </summary>
    /// <param name="registration">Handler registration.</param>
    /// <param name="replaceExisting">Whether conflicting custom handlers and built-in extensions may be replaced.</param>
    /// <returns>This builder.</returns>
    public OfficeDocumentReaderBuilder AddHandler(ReaderHandlerRegistration registration, bool replaceExisting = false) {
        _handlers.Register(registration, replaceExisting, preserveExistingCustomExtensions: false);
        return this;
    }

    /// <summary>
    /// Adds a handler while leaving extensions already owned by other custom handlers untouched.
    /// </summary>
    /// <param name="registration">Handler registration.</param>
    /// <param name="replaceExisting">Whether a handler with the same identifier and built-in extensions may be replaced.</param>
    /// <returns>This builder.</returns>
    public OfficeDocumentReaderBuilder AddHandlerPreservingExistingCustomExtensions(
        ReaderHandlerRegistration registration,
        bool replaceExisting = false) {
        _handlers.Register(registration, replaceExisting, preserveExistingCustomExtensions: true);
        return this;
    }

    /// <summary>
    /// Creates an immutable, thread-safe reader from the current configuration.
    /// </summary>
    public OfficeDocumentReader Build() {
        return new OfficeDocumentReader(_handlers.CaptureSnapshot());
    }
}
