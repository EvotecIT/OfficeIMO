using OfficeIMO.GoogleWorkspace;
using System.Net;
using System.Net.Http.Headers;
using System.Text;

namespace OfficeIMO.GoogleWorkspace.Drive {
    public sealed class GoogleDriveUploadOptions {
        public string Name { get; set; } = string.Empty;
        public string ContentType { get; set; } = "application/octet-stream";
        public string? ParentId { get; set; }
        public string? ConvertToGoogleMimeType { get; set; }
        public int ResumableChunkSize { get; set; } = 8 * 1024 * 1024;
        public IProgress<GoogleDriveTransferProgress>? Progress { get; set; }
    }

    public sealed class GoogleDriveTransferProgress {
        public GoogleDriveTransferProgress(long bytesTransferred, long? totalBytes) {
            BytesTransferred = bytesTransferred;
            TotalBytes = totalBytes;
        }

        public long BytesTransferred { get; }
        public long? TotalBytes { get; }
        public double? Percentage => TotalBytes > 0 ? (double)BytesTransferred / TotalBytes.Value * 100d : null;
    }

    public sealed partial class GoogleDriveClient {
        internal const long MultipartUploadLimitBytes = 5L * 1024 * 1024;

        public async Task<GoogleDriveFile> UploadMultipartAsync(
            byte[] content,
            GoogleDriveUploadOptions options,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            if (content == null) throw new ArgumentNullException(nameof(content));
            ValidateUploadOptions(options);
            if (content.LongLength > MultipartUploadLimitBytes) {
                throw new ArgumentOutOfRangeException(nameof(content), "Multipart uploads are limited to 5 MB. Use UploadResumableAsync for larger content.");
            }

            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.WriteScopes, report, "Google Drive multipart upload", cancellationToken).ConfigureAwait(false);
            string uri = $"https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives={Bool(Options.SupportsAllDrives)}&fields={Escape(DefaultFileFields)}";
            var file = await Transport.SendAsync<GoogleDriveFile>(
                token,
                HttpMethod.Post,
                uri,
                () => CreateMultipartContent(content, options),
                GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveFile,
                cancellationToken).ConfigureAwait(false);
            options.Progress?.Report(new GoogleDriveTransferProgress(content.LongLength, content.LongLength));
            return file;
        }

        public async Task<GoogleDriveFile> UploadResumableAsync(
            byte[] content,
            GoogleDriveUploadOptions options,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            if (content == null) throw new ArgumentNullException(nameof(content));
            ValidateUploadOptions(options);
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.WriteScopes, report, "Google Drive resumable upload", cancellationToken).ConfigureAwait(false);
            int chunkSize = NormalizeChunkSize(options.ResumableChunkSize);
            string metadataJson = SerializeUploadMetadata(options);
            string initUri = $"https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives={Bool(Options.SupportsAllDrives)}&fields={Escape(DefaultFileFields)}";
            var initiation = await Transport.SendRawAsync(
                token,
                HttpMethod.Post,
                initUri,
                () => new StringContent(metadataJson, Encoding.UTF8, "application/json"),
                GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Drive API",
                report,
                cancellationToken,
                request => {
                    request.Headers.TryAddWithoutValidation("X-Upload-Content-Type", options.ContentType);
                    request.Headers.TryAddWithoutValidation("X-Upload-Content-Length", content.LongLength.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }).ConfigureAwait(false);
            string sessionUri = initiation.GetHeader("Location")
                ?? throw new InvalidOperationException("Google Drive did not return a resumable upload session URI.");

            if (content.LongLength == 0) {
                GoogleWorkspaceHttpResponse response = await QueryResumableStatusAsync(
                    token,
                    sessionUri,
                    0,
                    report,
                    cancellationToken).ConfigureAwait(false);
                if (response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.Created) {
                    options.Progress?.Report(new GoogleDriveTransferProgress(0, 0));
                    return response.DeserializeJson(GoogleDriveJsonSerializerContext.Default.GoogleDriveFile);
                }

                throw new InvalidOperationException("Google Drive did not complete the zero-byte resumable upload.");
            }

            long offset = 0;
            while (offset < content.LongLength) {
                cancellationToken.ThrowIfCancellationRequested();
                int currentLength = (int)Math.Min(chunkSize, content.LongLength - offset);
                byte[] chunk = new byte[currentLength];
                Buffer.BlockCopy(content, (int)offset, chunk, 0, currentLength);
                long end = offset + currentLength - 1;

                try {
                    var response = await SendResumableChunkAsync(
                        token,
                        sessionUri,
                        chunk,
                        options.ContentType,
                        offset,
                        end,
                        content.LongLength,
                        report,
                        cancellationToken).ConfigureAwait(false);
                    if (response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.Created) {
                        options.Progress?.Report(new GoogleDriveTransferProgress(content.LongLength, content.LongLength));
                        return response.DeserializeJson(GoogleDriveJsonSerializerContext.Default.GoogleDriveFile);
                    }

                    offset = ResolveNextOffset(response, offset);
                    options.Progress?.Report(new GoogleDriveTransferProgress(offset, content.LongLength));
                } catch (Exception exception) when (IsAmbiguousResumableChunkFailure(exception, cancellationToken)) {
                    var status = await QueryResumableStatusAsync(token, sessionUri, content.LongLength, report, cancellationToken).ConfigureAwait(false);
                    if (status.StatusCode == HttpStatusCode.OK || status.StatusCode == HttpStatusCode.Created) {
                        options.Progress?.Report(new GoogleDriveTransferProgress(content.LongLength, content.LongLength));
                        return status.DeserializeJson(GoogleDriveJsonSerializerContext.Default.GoogleDriveFile);
                    }

                    offset = ResolveNextOffset(status, offset);
                    options.Progress?.Report(new GoogleDriveTransferProgress(offset, content.LongLength));
                }
            }

            throw new InvalidOperationException("Google Drive resumable upload ended without final file metadata.");
        }

        private static bool IsAmbiguousResumableChunkFailure(Exception exception, CancellationToken cancellationToken) {
            if (cancellationToken.IsCancellationRequested) return false;
            if (exception is GoogleWorkspaceApiException apiException) {
                return (int)apiException.ResponseStatusCode >= 500;
            }

            return exception is HttpRequestException || exception is TaskCanceledException;
        }

        public async Task<byte[]> DownloadAsync(
            string fileId,
            IProgress<GoogleDriveTransferProgress>? progress = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default,
            long? maxResponseBytes = null) {
            if (string.IsNullOrWhiteSpace(fileId)) throw new ArgumentException("File id is required.", nameof(fileId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.ReadScopes, report, "Google Drive file download", cancellationToken).ConfigureAwait(false);
            byte[] bytes = await Transport.SendBytesAsync(
                token,
                HttpMethod.Get,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}?alt=media&supportsAllDrives={Bool(Options.SupportsAllDrives)}",
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                cancellationToken,
                maxResponseBytes: maxResponseBytes ?? Options.MaxDownloadBytes).ConfigureAwait(false);
            progress?.Report(new GoogleDriveTransferProgress(bytes.LongLength, bytes.LongLength));
            return bytes;
        }

        public async Task<byte[]> ExportAsync(
            string fileId,
            string mimeType,
            IProgress<GoogleDriveTransferProgress>? progress = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default,
            long? maxResponseBytes = null) {
            if (string.IsNullOrWhiteSpace(fileId)) throw new ArgumentException("File id is required.", nameof(fileId));
            if (string.IsNullOrWhiteSpace(mimeType)) throw new ArgumentException("Export MIME type is required.", nameof(mimeType));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.ReadScopes, report, "Google Drive file export", cancellationToken).ConfigureAwait(false);
            byte[] bytes = await Transport.SendBytesAsync(
                token,
                HttpMethod.Get,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/export?mimeType={Escape(mimeType)}",
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                cancellationToken,
                maxResponseBytes: maxResponseBytes ?? Options.MaxDownloadBytes).ConfigureAwait(false);
            progress?.Report(new GoogleDriveTransferProgress(bytes.LongLength, bytes.LongLength));
            return bytes;
        }

        private Task<GoogleWorkspaceHttpResponse> SendResumableChunkAsync(
            string token,
            string sessionUri,
            byte[] chunk,
            string contentType,
            long start,
            long end,
            long total,
            TranslationReport report,
            CancellationToken cancellationToken) {
            return Transport.SendRawAsync(
                token,
                HttpMethod.Put,
                sessionUri,
                () => {
                    var content = new ByteArrayContent(chunk);
                    content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
                    content.Headers.TryAddWithoutValidation("Content-Range", $"bytes {start}-{end}/{total}");
                    return content;
                },
                GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Drive API",
                report,
                cancellationToken,
                additionalSuccessStatusCodes: new[] { (HttpStatusCode)308 },
                preserveRequestUri: true);
        }

        private Task<GoogleWorkspaceHttpResponse> QueryResumableStatusAsync(
            string token,
            string sessionUri,
            long total,
            TranslationReport report,
            CancellationToken cancellationToken) {
            return Transport.SendRawAsync(
                token,
                HttpMethod.Put,
                sessionUri,
                () => {
                    var content = new ByteArrayContent(Array.Empty<byte>());
                    content.Headers.TryAddWithoutValidation("Content-Range", $"bytes */{total}");
                    return content;
                },
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                cancellationToken,
                additionalSuccessStatusCodes: new[] { (HttpStatusCode)308 },
                preserveRequestUri: true);
        }

        private static MultipartContent CreateMultipartContent(byte[] content, GoogleDriveUploadOptions options) {
            string boundary = "officeimo-" + Guid.NewGuid().ToString("N");
            var multipart = new MultipartContent("related", boundary);
            multipart.Add(new StringContent(
                SerializeUploadMetadata(options),
                Encoding.UTF8,
                "application/json"));
            var media = new ByteArrayContent(content);
            media.Headers.ContentType = new MediaTypeHeaderValue(options.ContentType);
            multipart.Add(media);
            return multipart;
        }

        private static string SerializeUploadMetadata(GoogleDriveUploadOptions options) {
            return GoogleDriveJson.Serialize(new GoogleDriveFilePayload {
                Name = options.Name,
                MimeType = options.ConvertToGoogleMimeType,
                Parents = string.IsNullOrWhiteSpace(options.ParentId) ? null : new[] { options.ParentId! },
            }, GoogleDriveJsonSerializerContext.Default.GoogleDriveFilePayload);
        }

        private static long ResolveNextOffset(GoogleWorkspaceHttpResponse response, long fallback) {
            string? range = response.GetHeader("Range");
            if (string.IsNullOrWhiteSpace(range)) return fallback;
            int dash = range!.LastIndexOf('-');
            if (dash < 0 || !long.TryParse(range.Substring(dash + 1), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out long end)) {
                return fallback;
            }

            return end + 1;
        }

        private static int NormalizeChunkSize(int requested) {
            const int quantum = 256 * 1024;
            int bounded = Math.Max(quantum, requested);
            return bounded / quantum * quantum;
        }

        private static void ValidateUploadOptions(GoogleDriveUploadOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrWhiteSpace(options.Name)) throw new ArgumentException("Upload name is required.", nameof(options));
            if (string.IsNullOrWhiteSpace(options.ContentType)) throw new ArgumentException("Upload content type is required.", nameof(options));
        }
    }
}
