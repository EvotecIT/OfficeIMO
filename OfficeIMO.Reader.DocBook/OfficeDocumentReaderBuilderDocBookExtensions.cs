using System;
using OfficeIMO.Reader;

namespace OfficeIMO.Reader.DocBook;

/// <summary>Adds DocBook support to an isolated Reader builder.</summary>
public static class OfficeDocumentReaderBuilderDocBookExtensions {
    /// <summary>Stable handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.docbook";
    /// <summary>Adds .dbk and .docbook path and stream ingestion. Generic .xml remains owned by the XML adapter.</summary>
    public static OfficeDocumentReaderBuilder AddDocBookHandler(this OfficeDocumentReaderBuilder builder, ReaderDocBookOptions? options = null, bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderDocBookOptions registered = ReaderDocBookOptionsCloner.Clone(options);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO, Id = HandlerId, DisplayName = "DocBook Reader Adapter",
            Description = "Bounded DocBook adapter backed by the source-preserving OfficeIMO.DocBook engine.", Kind = ReaderInputKind.DocBook,
            Extensions = new[] { ".dbk", ".docbook" },
            DefaultMaxInputBytes = registered.ReadOptions.MaxInputBytes,
            ReadPath = (path, reader, token) => DocBookReaderAdapter.Read(path, reader, ReaderDocBookOptionsCloner.Clone(registered), token),
            ReadStream = (stream, name, reader, token) => DocBookReaderAdapter.Read(stream, name, reader, ReaderDocBookOptionsCloner.Clone(registered), token),
            ReadDocumentPath = (path, reader, token) => DocBookReaderAdapter.ReadDocument(path, reader, ReaderDocBookOptionsCloner.Clone(registered), token),
            ReadDocumentStream = (stream, name, reader, token) => DocBookReaderAdapter.ReadDocument(stream, name, reader, ReaderDocBookOptionsCloner.Clone(registered), token),
            WarningBehavior = ReaderWarningBehavior.Mixed, DeterministicOutput = true
        }, replaceExisting);
    }
}
