using System;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO;
using OfficeIMO.DocBook;
using OfficeIMO.Reader;

namespace OfficeIMO.Reader.DocBook;

internal static partial class DocBookReaderAdapter {
    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions? readerOptions = null, ReaderDocBookOptions? docBookOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, reader.MaxInputBytes);
        ReaderDocBookOptions adapter = ReaderDocBookOptionsCloner.Clone(docBookOptions);
        ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        cancellationToken.ThrowIfCancellationRequested();
        return BuildDocumentResult(DocBookDocument.Load(path, adapter.ReadOptions, cancellationToken), path, reader, adapter, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderDocBookOptions? docBookOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        ReaderOptions reader = readerOptions ?? new ReaderOptions();
        ReaderDocBookOptions adapter = ReaderDocBookOptionsCloner.Clone(docBookOptions);
        ApplyReaderLimit(adapter.ReadOptions, reader.MaxInputBytes);
        string name = string.IsNullOrWhiteSpace(sourceName) ? "document.xml" : sourceName!;
        return BuildDocumentResult(DocBookDocument.Load(stream, adapter.ReadOptions, cancellationToken), name, reader, adapter, cancellationToken);
    }

    private static OfficeDocumentReadResult BuildDocumentResult(DocBookDocument document, string sourceName, ReaderOptions reader, ReaderDocBookOptions options, CancellationToken cancellationToken) {
        DocBookProjection projection = CreateProjection(document, sourceName, reader, options, includeChunkWarnings: false, cancellationToken);
        OfficeDocumentModel model = projection.Model;
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            projection.Chunks,
            ReaderInputKind.DocBook,
            new OfficeDocumentSource { Path = sourceName, Title = model.Source.Title, Author = model.Source.Author },
            new[] { "officeimo.reader.docbook.rich-v5", "officeimo.docbook.common-structure" });
        result.Metadata = model.Metadata.Select(MapMetadata).ToArray();
        result.Links = model.Links.Select(link => MapLink(link, sourceName)).ToArray();
        result.Diagnostics = projection.Diagnostics.Select(diagnostic => MapDiagnostic(diagnostic, sourceName)).ToArray();
        return result;
    }

    private static OfficeDocumentMetadataEntry MapMetadata(OfficeDocumentModelMetadataEntry metadata) => new OfficeDocumentMetadataEntry {
        Id = metadata.Id,
        Category = metadata.Category,
        Name = metadata.Name,
        Value = metadata.Value,
        ValueType = metadata.ValueType,
        SourceObjectId = metadata.SourceObjectId,
        Location = metadata.Location == null ? null : MapLocation(metadata.Location, metadata.Location.Path ?? "memory"),
        Attributes = metadata.Attributes
    };

    private static OfficeDocumentLink MapLink(OfficeDocumentModelLink link, string sourceName) => new OfficeDocumentLink {
        Id = link.Id,
        Kind = link.Kind,
        Uri = link.Uri,
        DestinationName = link.DestinationName,
        DestinationPageNumber = link.DestinationPageNumber,
        DestinationMode = link.DestinationMode,
        NamedAction = link.NamedAction,
        RemoteFile = link.RemoteFile,
        RemoteDestinationName = link.RemoteDestinationName,
        RemoteDestinationPageNumber = link.RemoteDestinationPageNumber,
        Text = link.Text,
        Location = MapLocation(link.Location, sourceName)
    };

    private static OfficeDocumentDiagnostic MapDiagnostic(DocBookDiagnostic diagnostic, string sourceName) => new OfficeDocumentDiagnostic {
        Severity = diagnostic.Severity == DocBookDiagnosticSeverity.Error ? OfficeDocumentDiagnosticSeverity.Error
            : diagnostic.Severity == DocBookDiagnosticSeverity.Warning ? OfficeDocumentDiagnosticSeverity.Warning
            : OfficeDocumentDiagnosticSeverity.Information,
        Category = diagnostic.Code.StartsWith("DB01", StringComparison.Ordinal) ? OfficeDocumentDiagnosticCategory.Parsing : OfficeDocumentDiagnosticCategory.Content,
        Code = diagnostic.Code,
        Message = diagnostic.Message,
        Source = OfficeDocumentReaderBuilderDocBookExtensions.HandlerId,
        IsRecoverable = diagnostic.Severity != DocBookDiagnosticSeverity.Error,
        Location = new ReaderLocation { Path = sourceName, HeadingPath = diagnostic.Path }
    };
}
