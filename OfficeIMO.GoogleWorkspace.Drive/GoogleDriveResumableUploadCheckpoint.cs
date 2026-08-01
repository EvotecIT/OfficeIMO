using OfficeIMO.GoogleWorkspace;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.GoogleWorkspace.Drive {
    internal static class GoogleDriveResumableSessionUri {
        internal static string Validate(string value) {
            if (!Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) ||
                !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
                !string.IsNullOrEmpty(uri.UserInfo) ||
                !uri.IsDefaultPort && uri.Port != 443 ||
                !uri.Host.EndsWith(".googleapis.com", StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("The resumable upload checkpoint contains an untrusted session URI.");
            }
            return uri.AbsoluteUri;
        }
    }

    /// <summary>
    /// Durable resumable-upload state. Its serialized value contains the Google upload session URI and must be
    /// protected like a credential; diagnostics and <see cref="ToString"/> never reveal it.
    /// </summary>
    public sealed class GoogleDriveResumableUploadCheckpoint {
        private const int Magic = 0x5557494F;
        private const byte Version = 1;
        private GoogleDriveResumableUploadCheckpoint(string value, string sessionUri, string sourceFingerprint,
            string metadataFingerprint, long totalBytes, long confirmedBytes) {
            Value = value; SessionUri = sessionUri; SourceFingerprint = sourceFingerprint;
            MetadataFingerprint = metadataFingerprint; TotalBytes = totalBytes; ConfirmedBytes = confirmedBytes;
        }
        /// <summary>Opaque persistence value containing the sensitive upload-session URI; callers must encrypt it at rest.</summary>
        public string Value { get; }
        /// <summary>SHA-256 of the complete local source.</summary>
        public string SourceFingerprint { get; }
        /// <summary>SHA-256 of the upload metadata.</summary>
        public string MetadataFingerprint { get; }
        /// <summary>Declared source length.</summary>
        public long TotalBytes { get; }
        /// <summary>Bytes confirmed by Google.</summary>
        public long ConfirmedBytes { get; }
        internal string SessionUri { get; }

        /// <summary>Parses a checkpoint obtained from <see cref="Value"/>.</summary>
        public static GoogleDriveResumableUploadCheckpoint Parse(string value) {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A checkpoint is required.", nameof(value));
            if (value.Length > 64 * 1024) throw new InvalidDataException("The upload checkpoint is too large.");
            try {
                string normalized = value.Replace('-', '+').Replace('_', '/');
                if (normalized.Length % 4 == 2) normalized += "=="; else if (normalized.Length % 4 == 3) normalized += "=";
                byte[] bytes = Convert.FromBase64String(normalized);
                using var stream = new MemoryStream(bytes, writable: false);
                using var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true);
                if (reader.ReadInt32() != Magic || reader.ReadByte() != Version) throw new InvalidDataException("The upload checkpoint version is unsupported.");
                string uri = ReadString(reader, 16 * 1024); string source = ReadString(reader, 256);
                string metadata = ReadString(reader, 256); long total = reader.ReadInt64(); long confirmed = reader.ReadInt64();
                uri = GoogleDriveResumableSessionUri.Validate(uri);
                if (source.Length != 64 || metadata.Length != 64 ||
                    total < 0 || confirmed < 0 || confirmed > total || stream.Position != stream.Length) throw new InvalidDataException("The upload checkpoint payload is invalid.");
                return new GoogleDriveResumableUploadCheckpoint(value, uri, source, metadata, total, confirmed);
            } catch (FormatException exception) { throw new InvalidDataException("The upload checkpoint is not valid Base64.", exception); }
        }

        internal static GoogleDriveResumableUploadCheckpoint Create(string sessionUri, string sourceFingerprint,
            string metadataFingerprint, long totalBytes, long confirmedBytes) {
            sessionUri = GoogleDriveResumableSessionUri.Validate(sessionUri);
            using var stream = new MemoryStream();
            using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
                writer.Write(Magic); writer.Write(Version); WriteString(writer, sessionUri);
                WriteString(writer, sourceFingerprint); WriteString(writer, metadataFingerprint);
                writer.Write(totalBytes); writer.Write(confirmedBytes);
            }
            string value = Convert.ToBase64String(stream.ToArray()).TrimEnd('=').Replace('+', '-').Replace('/', '_');
            return new GoogleDriveResumableUploadCheckpoint(value, sessionUri, sourceFingerprint,
                metadataFingerprint, totalBytes, confirmedBytes);
        }
        internal GoogleDriveResumableUploadCheckpoint Advance(long confirmedBytes) =>
            Create(SessionUri, SourceFingerprint, MetadataFingerprint, TotalBytes, confirmedBytes);
        /// <inheritdoc />
        public override string ToString() => $"GoogleDriveResumableUploadCheckpoint({ConfirmedBytes}/{TotalBytes})";
        private static void WriteString(BinaryWriter writer, string value) { byte[] bytes = Encoding.UTF8.GetBytes(value); writer.Write(bytes.Length); writer.Write(bytes); }
        private static string ReadString(BinaryReader reader, int maximum) { int length = reader.ReadInt32(); if (length < 0 || length > maximum) throw new InvalidDataException("The checkpoint string length is invalid."); byte[] bytes = reader.ReadBytes(length); if (bytes.Length != length) throw new EndOfStreamException(); return Encoding.UTF8.GetString(bytes); }
    }

    /// <summary>Completed file metadata plus final durable transfer checkpoint.</summary>
    public sealed class GoogleDriveResumableUploadResult {
        internal GoogleDriveResumableUploadResult(GoogleDriveFile file, GoogleDriveResumableUploadCheckpoint checkpoint) { File = file; Checkpoint = checkpoint; }
        public GoogleDriveFile File { get; }
        public GoogleDriveResumableUploadCheckpoint Checkpoint { get; }
    }

    public sealed partial class GoogleDriveClient {
        /// <summary>Uploads a seekable stream and persists progress after initiation and every confirmed chunk.</summary>
        public async Task<GoogleDriveResumableUploadResult> UploadResumableStreamAsync(Stream content, long length,
            GoogleDriveUploadOptions options, GoogleDriveResumableUploadCheckpoint? checkpoint = null,
            Func<GoogleDriveResumableUploadCheckpoint, CancellationToken, Task>? checkpointSink = null,
            TranslationReport? report = null, CancellationToken cancellationToken = default) {
            if (content == null) throw new ArgumentNullException(nameof(content));
            if (!content.CanRead || !content.CanSeek) throw new ArgumentException("A readable, seekable stream is required for restart-safe upload.", nameof(content));
            if (length < 0 || length > content.Length) throw new ArgumentOutOfRangeException(nameof(length));
            ValidateUploadOptions(options); report ??= new TranslationReport();
            string sourceFingerprint = ComputeFingerprint(content, length, cancellationToken);
            string metadataJson = SerializeUploadMetadata(options);
            string metadataFingerprint = ComputeFingerprint(Encoding.UTF8.GetBytes(metadataJson));
            string token = await AcquireTokenAsync(Options.WriteScopes, report, "Google Drive durable resumable upload", cancellationToken).ConfigureAwait(false);
            GoogleDriveResumableUploadCheckpoint state;
            if (checkpoint == null) {
                string initUri = $"https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives={Bool(Options.SupportsAllDrives)}&fields={Escape(DefaultFileFields)}";
                GoogleWorkspaceHttpResponse initiation = await Transport.SendRawAsync(token, HttpMethod.Post, initUri,
                    () => new StringContent(metadataJson, Encoding.UTF8, "application/json"),
                    GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API", report, cancellationToken,
                    request => { request.Headers.TryAddWithoutValidation("X-Upload-Content-Type", options.ContentType); request.Headers.TryAddWithoutValidation("X-Upload-Content-Length", length.ToString(System.Globalization.CultureInfo.InvariantCulture)); },
                    mutationKind: GoogleWorkspaceMutationKind.Action,
                    revisionPrecondition: GoogleWorkspaceRevisionPrecondition.ResumableSessionState(
                        CreateResumableInitiationState(metadataJson, length)),
                    requiredScopes: Options.WriteScopes).ConfigureAwait(false);
                string sessionUri = initiation.GetHeader("Location") ?? throw new InvalidOperationException("Google Drive did not return a resumable upload session URI.");
                state = GoogleDriveResumableUploadCheckpoint.Create(sessionUri, sourceFingerprint, metadataFingerprint, length, 0);
                if (checkpointSink != null) await checkpointSink(state, cancellationToken).ConfigureAwait(false);
            } else {
                if (!StringComparer.Ordinal.Equals(checkpoint.SourceFingerprint, sourceFingerprint) ||
                    !StringComparer.Ordinal.Equals(checkpoint.MetadataFingerprint, metadataFingerprint) || checkpoint.TotalBytes != length) {
                    throw new InvalidOperationException("The upload checkpoint belongs to changed content or upload metadata.");
                }
                state = checkpoint;
            }

            GoogleWorkspaceHttpTransport.DeferredMutation create = BeginResumableFileCreate(state.SessionUri);
            try {
                GoogleWorkspaceHttpResponse status = await QueryResumableStatusAsync(token, state.SessionUri, length, report, cancellationToken).ConfigureAwait(false);
                if (status.StatusCode == HttpStatusCode.OK || status.StatusCode == HttpStatusCode.Created) {
                    create.CompleteSuccess();
                    EnsureSourceUnchanged(content, length, sourceFingerprint, cancellationToken);
                    state = state.Advance(length); if (checkpointSink != null) await checkpointSink(state, cancellationToken).ConfigureAwait(false);
                    return new GoogleDriveResumableUploadResult(status.DeserializeJson(GoogleDriveJsonSerializerContext.Default.GoogleDriveFile), state);
                }
                long offset = ResolveNextOffset(status, state.ConfirmedBytes);
                int chunkSize = NormalizeChunkSize(options.ResumableChunkSize);
                var buffer = new byte[chunkSize];
                int noProgressResponses = 0;
                while (offset < length) {
                    cancellationToken.ThrowIfCancellationRequested(); content.Position = offset;
                    int wanted = (int)Math.Min(buffer.Length, length - offset); int read = 0;
                    while (read < wanted) { int current = await content.ReadAsync(buffer, read, wanted - read, cancellationToken).ConfigureAwait(false); if (current == 0) throw new EndOfStreamException("The upload source changed during transfer."); read += current; }
                    byte[] chunk = read == buffer.Length ? buffer : buffer.Take(read).ToArray();
                    try {
                        status = await SendResumableChunkAsync(token, state.SessionUri, chunk, options.ContentType,
                            offset, offset + read - 1, length, report, cancellationToken).ConfigureAwait(false);
                    } catch (Exception exception) when (IsAmbiguousResumableChunkFailure(exception, cancellationToken)) {
                        status = await QueryResumableStatusAsync(token, state.SessionUri, length, report, cancellationToken).ConfigureAwait(false);
                    }
                    bool completed = status.StatusCode == HttpStatusCode.OK || status.StatusCode == HttpStatusCode.Created;
                    if (completed) offset = length;
                    else {
                        long nextOffset = ResolveNextOffset(status, offset);
                        if (nextOffset <= offset) {
                            if (++noProgressResponses > _session.Options.MaxRetryCount) throw new InvalidDataException("Google Drive repeatedly failed to confirm progress for the resumable upload chunk.");
                        } else {
                            offset = nextOffset;
                            noProgressResponses = 0;
                        }
                    }
                    if (completed) create.CompleteSuccess();
                    state = state.Advance(offset); if (checkpointSink != null) await checkpointSink(state, cancellationToken).ConfigureAwait(false);
                    options.Progress?.Report(new GoogleDriveTransferProgress(offset, length));
                    if (offset == length && completed) {
                        EnsureSourceUnchanged(content, length, sourceFingerprint, cancellationToken);
                        return new GoogleDriveResumableUploadResult(status.DeserializeJson(GoogleDriveJsonSerializerContext.Default.GoogleDriveFile), state);
                    }
                }
                throw new InvalidOperationException("Google Drive resumable upload ended without final file metadata.");
            } catch (Exception exception) {
                create.CompleteFailure(exception);
                throw;
            }
        }

        /// <summary>Uploads a file through the durable resumable path.</summary>
        public async Task<GoogleDriveResumableUploadResult> UploadResumableFileAsync(string path,
            GoogleDriveUploadOptions options, GoogleDriveResumableUploadCheckpoint? checkpoint = null,
            Func<GoogleDriveResumableUploadCheckpoint, CancellationToken, Task>? checkpointSink = null,
            TranslationReport? report = null, CancellationToken cancellationToken = default) {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 64 * 1024, FileOptions.SequentialScan);
            return await UploadResumableStreamAsync(stream, stream.Length, options, checkpoint, checkpointSink, report, cancellationToken).ConfigureAwait(false);
        }

        private static string ComputeFingerprint(Stream stream, long length, CancellationToken cancellationToken) {
            long position = stream.Position; try { stream.Position = 0; using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256); var buffer = new byte[64 * 1024]; long remaining = length; while (remaining > 0) { cancellationToken.ThrowIfCancellationRequested(); int read = stream.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining)); if (read == 0) throw new EndOfStreamException(); hash.AppendData(buffer, 0, read); remaining -= read; } return ToHex(hash.GetHashAndReset()); } finally { stream.Position = position; }
        }
        private static string ComputeFingerprint(byte[] bytes) { using var hash = SHA256.Create(); return ToHex(hash.ComputeHash(bytes)); }
        private static void EnsureSourceUnchanged(Stream stream, long length, string expected,
            CancellationToken cancellationToken) {
            if (!StringComparer.Ordinal.Equals(expected, ComputeFingerprint(stream, length, cancellationToken))) {
                throw new InvalidOperationException("The upload source changed during transfer; the remote outcome must be reconciled before retrying.");
            }
        }
        private static string ToHex(byte[] bytes) => BitConverter.ToString(bytes).Replace("-", string.Empty).ToLowerInvariant();
    }
}
