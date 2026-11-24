using OfficeIMO.Core.Internal;

namespace OfficeIMO.Notebook;

public sealed partial class NotebookDocument {
    private const int FlushThreshold = 128 * 1024;

    /// <summary>Saves to the path supplied to Create or Load.</summary>
    public void Save(NotebookSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        Save(RequirePath(), options, cancellationToken);

    /// <summary>Writes a notebook file and publishes it after serialization succeeds.</summary>
    public void Save(string path, NotebookSaveOptions? options = null, CancellationToken cancellationToken = default) {
        string destination = GetPath(path);
        NotebookSaveOptions effective = PrepareSave(options, cancellationToken);
        OfficeFileCommit.WriteAtomically(destination, output => Write(output, effective, cancellationToken), cancellationToken);
    }

    /// <summary>Writes an nbformat 4 notebook at the stream's current position, leaving the stream open.</summary>
    public void Save(Stream stream, NotebookSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        Write(stream, PrepareSave(options, cancellationToken), cancellationToken);

    /// <summary>Asynchronously saves to the path supplied to Create or Load.</summary>
    public Task SaveAsync(NotebookSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        SaveAsync(RequirePath(), options, cancellationToken);

    /// <summary>Asynchronously writes a notebook file and publishes it after serialization succeeds.</summary>
    public async Task SaveAsync(string path, NotebookSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        string destination = GetPath(path);
        NotebookSaveOptions effective = PrepareSave(options, cancellationToken);
        await OfficeFileCommit.WriteAtomicallyAsync(destination, (output, token) => WriteAsync(output, effective, token),
            cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes a notebook at the stream's current position, leaving it open.</summary>
    public Task SaveAsync(Stream stream, NotebookSaveOptions? options = null, CancellationToken cancellationToken = default) =>
        WriteAsync(stream, PrepareSave(options, cancellationToken), cancellationToken);

    /// <summary>Returns the serialized notebook as UTF-8 bytes.</summary>
    public byte[] ToBytes(NotebookSaveOptions? options = null, CancellationToken cancellationToken = default) {
        using var destination = new MemoryStream();
        Save(destination, options, cancellationToken);
        return destination.ToArray();
    }

    private NotebookSaveOptions PrepareSave(NotebookSaveOptions? options, CancellationToken cancellationToken) {
        CheckDisposed();
        cancellationToken.ThrowIfCancellationRequested();
        var effective = new NotebookSaveOptions { Indented = options?.Indented ?? true, AllowLoss = options?.AllowLoss ?? false };
        if (!effective.AllowLoss) new OfficeConversionReport("Jupyter Notebook", "Jupyter Notebook", diagnostics: Diagnostics).ThrowIfLoss();
        return effective;
    }

    private void Write(Stream stream, NotebookSaveOptions options, CancellationToken cancellationToken) {
        using Utf8JsonWriter writer = OpenWriter(stream, options);
        try {
            WriteHeader(writer);
            foreach (NotebookCell cell in _cells) {
                cancellationToken.ThrowIfCancellationRequested();
                NotebookWriter.WriteCell(writer, cell, NbFormat == 3, cancellationToken);
                if (writer.BytesPending >= FlushThreshold) writer.Flush();
            }
            WriteFooter(writer);
            cancellationToken.ThrowIfCancellationRequested();
            writer.Flush();
        } finally {
            // Keep disposal from retrying a failed or canceled flush against the destination.
            writer.Reset(Stream.Null);
        }
    }

    private async Task WriteAsync(Stream stream, NotebookSaveOptions options, CancellationToken cancellationToken) {
        using Utf8JsonWriter writer = OpenWriter(stream, options);
        try {
            WriteHeader(writer);
            foreach (NotebookCell cell in _cells) {
                cancellationToken.ThrowIfCancellationRequested();
                NotebookWriter.WriteCell(writer, cell, NbFormat == 3, cancellationToken);
                if (writer.BytesPending >= FlushThreshold) await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
            WriteFooter(writer);
            await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
        } finally {
            // Disposal flushes its target synchronously. Detach the stream and discard any unwritten data.
            writer.Reset(Stream.Null);
        }
    }

    private static Utf8JsonWriter OpenWriter(Stream stream, NotebookSaveOptions options) {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanWrite) throw new ArgumentException("The notebook destination must be writable.", nameof(stream));
        return new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = options.Indented });
    }

    /// <summary>Gets whether the native writer retains opaque root properties that other formats may omit.</summary>
    internal bool HasAdditionalSourceProperties => _json != null && Root.EnumerateObject().Any(IsAdditionalSourceProperty);

    private bool IsAdditionalSourceProperty(JsonProperty property) =>
        property.Name is not ("nbformat" or "nbformat_minor" or "metadata" or "cells") &&
        !(NbFormat == 3 && property.NameEquals("worksheets"));

    private void WriteHeader(Utf8JsonWriter writer) {
        writer.WriteStartObject();
        NotebookJson.CopyProperties(writer, Root, IsAdditionalSourceProperty);
        writer.WriteNumber("nbformat", 4);
        writer.WriteNumber("nbformat_minor", NbFormat == 3 ? CurrentMinor : Math.Max(CurrentMinor, NbFormatMinor));
        if (_metadataEdited || !NotebookJson.WriteOriginal(writer, "metadata", NotebookJson.Property(Root, "metadata"))) {
            NotebookJson.WriteMap(writer, "metadata", Metadata);
        }
        writer.WriteStartArray("cells");
    }

    private static void WriteFooter(Utf8JsonWriter writer) {
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    private string RequirePath() => _path ?? throw new InvalidOperationException("This notebook has no default save destination.");
}
