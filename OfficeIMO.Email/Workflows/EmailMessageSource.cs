using OfficeIMO.Email.Data;

namespace OfficeIMO.Email;

// A local-file lease locator, not a durable checkpoint. Timestamp-preserving edits are unsupported.
internal sealed class EmailMessageSource {
    private readonly long _length;
    private readonly DateTime _lastWrite;
    internal EmailMessageSource(string path) {
        Path = System.IO.Path.GetFullPath(path);
        var info = new FileInfo(Path);
        if (!info.Exists) throw new FileNotFoundException("An existing local email or store file is required.", Path);
        _length = info.Length;
        _lastWrite = info.LastWriteTimeUtc;
    }
    internal string Path { get; }
    internal void Validate() {
        var info = new FileInfo(Path);
        if (!info.Exists || info.Length != _length || info.LastWriteTimeUtc != _lastWrite)
            throw new IOException("The email source changed or was removed. Read it again before saving attachments or exporting.");
    }
    internal EmailDataOpenResult Open(bool content, CancellationToken token) {
        EmailDataOpenResult data = EmailDataArtifact.Open(Path,
            new EmailDataOpenOptions(email: new EmailReaderOptions(includeAttachmentContent: content,
                includeEmbeddedMessages: content), useStreamingEmailReader: content), token);
        if (data.Email?.HasErrors == true) {
            string errors = string.Join("; ", data.Email.Diagnostics.Where(d => d.Severity == EmailDiagnosticSeverity.Error)
                .Select(d => d.Code + ": " + d.Message));
            data.Dispose();
            throw new InvalidDataException(errors);
        }
        return data;
    }
}
