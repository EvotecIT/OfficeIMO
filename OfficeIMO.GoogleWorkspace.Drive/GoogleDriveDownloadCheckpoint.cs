using OfficeIMO.GoogleWorkspace;
using System.IO;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.GoogleWorkspace.Drive {
    /// <summary>Portable file-download checkpoint bound to resource, revision, destination, length, and content prefix.</summary>
    public sealed class GoogleDriveDownloadCheckpoint {
        private const int Magic = 0x4447494F;
        private const byte Version = 2;
        private GoogleDriveDownloadCheckpoint(string value, string fileId, long expectedVersion, long totalBytes,
            long confirmedBytes, int chunkSize, string destinationIdentity, string prefixFingerprint) {
            Value = value; FileId = fileId; ExpectedVersion = expectedVersion; TotalBytes = totalBytes;
            ConfirmedBytes = confirmedBytes; ChunkSize = chunkSize; DestinationIdentity = destinationIdentity; PrefixFingerprint = prefixFingerprint;
        }
        public string Value { get; }
        public string FileId { get; }
        public long ExpectedVersion { get; }
        public long TotalBytes { get; }
        public long ConfirmedBytes { get; }
        public int ChunkSize { get; }
        public string DestinationIdentity { get; }
        public string PrefixFingerprint { get; }

        public static GoogleDriveDownloadCheckpoint Parse(string value) {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A checkpoint is required.", nameof(value));
            if (value.Length > 64 * 1024) throw new InvalidDataException("The download checkpoint is too large.");
            try {
                string normalized = value.Replace('-', '+').Replace('_', '/'); if (normalized.Length % 4 == 2) normalized += "=="; else if (normalized.Length % 4 == 3) normalized += "=";
                byte[] bytes = Convert.FromBase64String(normalized); using var stream = new MemoryStream(bytes, false); using var reader = new BinaryReader(stream, Encoding.UTF8, true);
                if (reader.ReadInt32() != Magic || reader.ReadByte() != Version) throw new InvalidDataException("The download checkpoint version is unsupported.");
                string fileId = ReadString(reader, 2048); long revision = reader.ReadInt64(); long total = reader.ReadInt64(); long confirmed = reader.ReadInt64(); int chunkSize = reader.ReadInt32();
                string destination = ReadString(reader, 256); string prefix = ReadString(reader, 256);
                if (revision < 0 || total < 0 || confirmed < 0 || confirmed > total || chunkSize < 256 * 1024 || destination.Length != 64 || prefix.Length != 64 || stream.Position != stream.Length) throw new InvalidDataException("The download checkpoint payload is invalid.");
                return new GoogleDriveDownloadCheckpoint(value, fileId, revision, total, confirmed, chunkSize, destination, prefix);
            } catch (FormatException exception) { throw new InvalidDataException("The download checkpoint is not valid Base64.", exception); }
        }
        internal static GoogleDriveDownloadCheckpoint Create(string fileId, long expectedVersion, long totalBytes,
            long confirmedBytes, int chunkSize, string destinationIdentity, string prefixFingerprint) {
            using var stream = new MemoryStream(); using (var writer = new BinaryWriter(stream, Encoding.UTF8, true)) {
                writer.Write(Magic); writer.Write(Version); WriteString(writer, fileId); writer.Write(expectedVersion);
                writer.Write(totalBytes); writer.Write(confirmedBytes); writer.Write(chunkSize); WriteString(writer, destinationIdentity); WriteString(writer, prefixFingerprint);
            }
            string value = Convert.ToBase64String(stream.ToArray()).TrimEnd('=').Replace('+', '-').Replace('/', '_');
            return new GoogleDriveDownloadCheckpoint(value, fileId, expectedVersion, totalBytes, confirmedBytes, chunkSize, destinationIdentity, prefixFingerprint);
        }
        public override string ToString() => $"GoogleDriveDownloadCheckpoint({ConfirmedBytes}/{TotalBytes})";
        private static void WriteString(BinaryWriter writer, string value) { byte[] bytes = Encoding.UTF8.GetBytes(value); writer.Write(bytes.Length); writer.Write(bytes); }
        private static string ReadString(BinaryReader reader, int maximum) { int length = reader.ReadInt32(); if (length < 0 || length > maximum) throw new InvalidDataException("The checkpoint string length is invalid."); byte[] bytes = reader.ReadBytes(length); if (bytes.Length != length) throw new EndOfStreamException(); return Encoding.UTF8.GetString(bytes); }
    }

    public sealed partial class GoogleDriveClient {
        /// <summary>Downloads to a new file or resumes an exact checkpoint after process restart.</summary>
        public async Task<GoogleDriveDownloadCheckpoint> DownloadToFileAsync(string fileId, string destinationPath,
            GoogleDriveDownloadCheckpoint? checkpoint = null,
            Func<GoogleDriveDownloadCheckpoint, CancellationToken, Task>? checkpointSink = null,
            int chunkSize = 8 * 1024 * 1024, TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(fileId)) throw new ArgumentException("File id is required.", nameof(fileId));
            if (string.IsNullOrWhiteSpace(destinationPath)) throw new ArgumentException("Destination path is required.", nameof(destinationPath));
            if (chunkSize < 256 * 1024) throw new ArgumentOutOfRangeException(nameof(chunkSize));
            string fullPath = Path.GetFullPath(destinationPath); string destinationIdentity = HashText(fullPath);
            report ??= new TranslationReport(); GoogleDriveFile metadata = await GetFileAsync(fileId, DefaultFileFields, report, cancellationToken).ConfigureAwait(false);
            long version = metadata.Version ?? throw new InvalidOperationException("Google Drive did not provide a file version for guarded download.");
            long total = metadata.Size ?? throw new InvalidOperationException("Google Drive did not provide a file size for guarded download.");
            if (total < 0) {
                throw new InvalidDataException($"Google Drive declared an invalid negative file size of {total} bytes.");
            }
            if (total > Options.MaxDownloadBytes) {
                throw new InvalidDataException(
                    $"Google Drive declared {total} bytes, exceeding the configured download limit of {Options.MaxDownloadBytes} bytes.");
            }
            string token = await AcquireTokenAsync(Options.ReadScopes, report,
                "Google Drive durable file download", cancellationToken).ConfigureAwait(false);
            long offset; string contentFingerprint;
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
            using FileStream output = checkpoint == null
                ? GoogleDriveDownloadFileGuard.CreateNew(fullPath, 64 * 1024)
                : GoogleDriveDownloadFileGuard.OpenExisting(fullPath, 64 * 1024);
            if (checkpoint == null) {
                offset = 0; contentFingerprint = EmptyContentFingerprint();
            } else {
                if (!StringComparer.Ordinal.Equals(checkpoint.FileId, fileId) || checkpoint.ExpectedVersion != version || checkpoint.TotalBytes != total || checkpoint.ChunkSize != chunkSize || !StringComparer.Ordinal.Equals(checkpoint.DestinationIdentity, destinationIdentity)) throw new InvalidOperationException("The download checkpoint belongs to a different resource, revision, chunking policy, or destination.");
                if (checkpoint.ConfirmedBytes == 0 && output.Length != 0) {
                    throw new InvalidOperationException("The zero-byte download checkpoint destination is no longer empty.");
                }
                if (output.Length < checkpoint.ConfirmedBytes ||
                    !StringComparer.Ordinal.Equals(checkpoint.PrefixFingerprint,
                        HashFileChain(output, chunkSize, checkpoint.ConfirmedBytes, cancellationToken))) {
                    throw new InvalidOperationException("The checkpointed destination prefix changed after the checkpoint was saved.");
                }
                if (output.Length > checkpoint.ConfirmedBytes) {
                    output.SetLength(checkpoint.ConfirmedBytes);
                    output.Flush(flushToDisk: true);
                }
                offset = checkpoint.ConfirmedBytes; contentFingerprint = checkpoint.PrefixFingerprint;
            }
            output.Position = offset;
            GoogleDriveDownloadCheckpoint state = GoogleDriveDownloadCheckpoint.Create(fileId, version, total, offset, chunkSize, destinationIdentity, contentFingerprint);
            try {
                await PersistDownloadCheckpointAsync(fullPath, output, state, checkpointSink, cancellationToken).ConfigureAwait(false);
            } catch (Exception exception) when (checkpoint == null) {
                RemoveUncheckpointedDestination(fullPath, output, state, exception);
                throw;
            }
            while (offset < total) {
                cancellationToken.ThrowIfCancellationRequested(); long end = Math.Min(total - 1, offset + chunkSize - 1); long expected = end - offset + 1;
                byte[] bytes = await Transport.SendBytesAsync(token, HttpMethod.Get,
                    $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}?alt=media&supportsAllDrives={Bool(Options.SupportsAllDrives)}",
                    GoogleWorkspaceRequestSafety.Safe, "Google Drive API", report, cancellationToken,
                    maxResponseBytes: expected,
                    configureRequest: request => request.Headers.Range = new RangeHeaderValue(offset, end),
                    validateResponse: response => ValidateRangedDownloadResponse(response, offset, end, total)).ConfigureAwait(false);
                if (bytes.LongLength != expected) throw new InvalidDataException("Google Drive returned an unexpected ranged-download length.");
                GoogleDriveDownloadFileGuard.EnsurePathReferencesHandle(fullPath, output);
                await output.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
                await output.FlushAsync(cancellationToken).ConfigureAwait(false);
                output.Flush(flushToDisk: true);
                offset += bytes.LongLength;
                contentFingerprint = ExtendContentFingerprint(contentFingerprint, bytes, bytes.Length);
                state = GoogleDriveDownloadCheckpoint.Create(fileId, version, total, offset, chunkSize, destinationIdentity, contentFingerprint);
                await PersistDownloadCheckpointAsync(fullPath, output, state, checkpointSink, cancellationToken).ConfigureAwait(false);
            }
            GoogleDriveFile finalMetadata = await GetFileAsync(fileId, DefaultFileFields, report, cancellationToken).ConfigureAwait(false);
            if (finalMetadata.Version != version || finalMetadata.Size != total) throw new InvalidOperationException("The Google Drive resource changed during download; the destination was retained for reconciliation.");
            GoogleDriveDownloadFileGuard.EnsurePathReferencesHandle(fullPath, output);
            if (output.Length != state.ConfirmedBytes ||
                !StringComparer.Ordinal.Equals(state.PrefixFingerprint,
                    HashFileChain(output, chunkSize, state.ConfirmedBytes, cancellationToken)) ||
                output.Length != state.ConfirmedBytes) {
                throw new InvalidOperationException("The guarded download destination changed before completion; the checkpoint and destination must be reconciled.");
            }
            GoogleDriveDownloadFileGuard.EnsurePathReferencesHandle(fullPath, output);
            return state;
        }

        private static string HashText(string value) { using var hash = SHA256.Create(); return Hex(hash.ComputeHash(Encoding.UTF8.GetBytes(value))); }
        private static void ValidateRangedDownloadResponse(HttpResponseMessage response, long start, long end, long total) {
            ContentRangeHeaderValue? range = response.Content.Headers.ContentRange;
            if (response.StatusCode != System.Net.HttpStatusCode.PartialContent
                || range == null
                || !StringComparer.OrdinalIgnoreCase.Equals(range.Unit, "bytes")
                || range.From != start
                || range.To != end
                || range.Length != total) {
                throw new InvalidDataException(
                    $"Google Drive did not confirm the requested byte range {start}-{end}/{total}.");
            }
        }
        private static string EmptyContentFingerprint() => HashText("OfficeIMO.GoogleDriveDownloadCheckpoint.v2");
        private static string HashFileChain(Stream stream, int chunkSize, long length, CancellationToken cancellationToken) {
            string current = EmptyContentFingerprint();
            long position = stream.Position;
            stream.Position = 0;
            var buffer = new byte[chunkSize]; long remaining = length;
            try {
                while (remaining > 0) {
                    int read = ReadChunk(stream, buffer, (int)Math.Min(buffer.Length, remaining), cancellationToken);
                    if (read == 0) throw new EndOfStreamException("The partial download is shorter than its checkpoint.");
                    current = ExtendContentFingerprint(current, buffer, read);
                    remaining -= read;
                }
                return current;
            } finally {
                stream.Position = position;
            }
        }

        private static async Task PersistDownloadCheckpointAsync(string path, FileStream stream,
            GoogleDriveDownloadCheckpoint checkpoint,
            Func<GoogleDriveDownloadCheckpoint, CancellationToken, Task>? checkpointSink,
            CancellationToken cancellationToken) {
            GoogleDriveDownloadFileGuard.EnsurePathReferencesHandle(path, stream);
            if (checkpointSink != null) await checkpointSink(checkpoint, cancellationToken).ConfigureAwait(false);
            GoogleDriveDownloadFileGuard.EnsurePathReferencesHandle(path, stream);
        }
        private static void RemoveUncheckpointedDestination(string path, FileStream stream,
            GoogleDriveDownloadCheckpoint checkpoint, Exception checkpointException) {
            try {
                GoogleDriveDownloadFileGuard.EnsurePathReferencesHandle(path, stream);
                if (stream.Length != 0) {
                    checkpointException.Data["OfficeIMO.GoogleWorkspace.UnpersistedDownloadCheckpoint"] = checkpoint.Value;
                    return;
                }

                stream.Dispose();
                File.Delete(path);
            } catch (Exception cleanupException) {
                checkpointException.Data["OfficeIMO.GoogleWorkspace.UnpersistedDownloadCheckpoint"] = checkpoint.Value;
                checkpointException.Data["OfficeIMO.GoogleWorkspace.DownloadCleanupFailure"] =
                    cleanupException.GetType().FullName ?? cleanupException.GetType().Name;
            }
        }
        private static int ReadChunk(Stream stream, byte[] buffer, int wanted, CancellationToken cancellationToken) {
            int total = 0;
            while (total < wanted) { cancellationToken.ThrowIfCancellationRequested(); int read = stream.Read(buffer, total, wanted - total); if (read == 0) break; total += read; }
            return total;
        }
        private static string ExtendContentFingerprint(string current, byte[] bytes, int count) {
            using var incremental = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            incremental.AppendData(ParseHex(current)); incremental.AppendData(bytes, 0, count);
            return Hex(incremental.GetHashAndReset());
        }
        private static byte[] ParseHex(string value) {
            var bytes = new byte[value.Length / 2];
            for (int index = 0; index < bytes.Length; index++) bytes[index] = Convert.ToByte(value.Substring(index * 2, 2), 16);
            return bytes;
        }
        private static string Hex(byte[] bytes) => BitConverter.ToString(bytes).Replace("-", string.Empty).ToLowerInvariant();
    }
}
