using System;
using OfficeIMO.Reader;

namespace OfficeIMO.Reader.Opml;

/// <summary>Adds OPML support to an isolated Reader builder.</summary>
public static class OfficeDocumentReaderBuilderOpmlExtensions {
    /// <summary>Stable handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.opml";
    /// <summary>Adds .opml path and stream ingestion.</summary>
    public static OfficeDocumentReaderBuilder AddOpmlHandler(this OfficeDocumentReaderBuilder builder, ReaderOpmlOptions? options = null, bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderOpmlOptions registered = ReaderOpmlOptionsCloner.Clone(options);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO, Id = HandlerId, DisplayName = "OPML Reader Adapter",
            Description = "Bounded OPML adapter backed by the lossless OfficeIMO.Opml engine.", Kind = ReaderInputKind.Opml,
            Extensions = new[] { ".opml" },
            DefaultMaxInputBytes = registered.ReadOptions.MaxInputBytes,
            ReadPath = (path, reader, token) => OpmlReaderAdapter.Read(path, reader, ReaderOpmlOptionsCloner.Clone(registered), token),
            ReadStream = (stream, name, reader, token) => OpmlReaderAdapter.Read(stream, name, reader, ReaderOpmlOptionsCloner.Clone(registered), token),
            WarningBehavior = ReaderWarningBehavior.WarningChunksOnly, DeterministicOutput = true
        }, replaceExisting);
    }
}
