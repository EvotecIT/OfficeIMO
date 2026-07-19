using System;

namespace OfficeIMO.Reader;

/// <summary>
/// Builds an <see cref="OfficeDocumentReader"/> with an isolated handler configuration.
/// </summary>
/// <remarks>
/// A builder may be reused or changed after <see cref="Build"/>. Each built reader retains its own
/// immutable snapshot and is unaffected by later builder changes.
/// </remarks>
public sealed partial class OfficeDocumentReaderBuilder {
    private readonly ReaderHandlerRegistry _handlers = new ReaderHandlerRegistry();
    private int _maxConcurrentReads = DocumentReaderEngine.DefaultMaxConcurrentReads;

    /// <summary>
    /// Adds a handler to this reader configuration.
    /// </summary>
    /// <param name="registration">Handler registration.</param>
    /// <param name="replaceExisting">Whether conflicting handler identifiers and extensions may be replaced.</param>
    /// <returns>This builder.</returns>
    public OfficeDocumentReaderBuilder AddHandler(ReaderHandlerRegistration registration, bool replaceExisting = false) {
        _handlers.Register(registration, replaceExisting);
        return this;
    }

    /// <summary>
    /// Sets the maximum number of asynchronous read operations allowed in flight for the built reader.
    /// </summary>
    /// <param name="maxConcurrentReads">A value from 1 through 64.</param>
    /// <returns>This builder.</returns>
    public OfficeDocumentReaderBuilder WithMaxConcurrentReads(int maxConcurrentReads) {
        if (maxConcurrentReads < 1 || maxConcurrentReads > DocumentReaderEngine.MaximumConcurrentReads) {
            throw new ArgumentOutOfRangeException(
                nameof(maxConcurrentReads),
                $"Max concurrent reads must be between 1 and {DocumentReaderEngine.MaximumConcurrentReads}.");
        }

        _maxConcurrentReads = maxConcurrentReads;
        return this;
    }

    /// <summary>
    /// Creates an immutable, thread-safe reader from the current configuration.
    /// </summary>
    public OfficeDocumentReader Build() {
        return new OfficeDocumentReader(
            _handlers.CaptureSnapshot(),
            _maxConcurrentReads,
            _processorPipelineBuilder.Build(),
            _processingOptions.Clone());
    }
}
