using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static async Task<SourceInfo> BuildSourceInfoFromStreamAsync(Stream stream, string? sourceName,
        bool computeHash, CancellationToken cancellationToken) {
        SourceInfo source = BuildSourceInfoFromStream(stream, sourceName, false, cancellationToken);
        if (computeHash) source.SourceHash = await TryComputeStreamSha256Async(stream, cancellationToken).ConfigureAwait(false);
        return source;
    }

    private static async Task<SourceInfo> BuildSourceInfoFromPathAsync(string path, bool computeHash,
        CancellationToken cancellationToken) {
        SourceInfo source = BuildSourceInfoFromPath(path, false, cancellationToken);
        if (computeHash) {
            try {
                using var stream = OpenAsyncReadStream(path);
                source.SourceHash = await TryComputeStreamSha256Async(stream, cancellationToken).ConfigureAwait(false);
            } catch (OperationCanceledException) { throw; }
            catch { /* Source metadata is best effort, matching synchronous reads. */ }
        }
        return source;
    }

    private static async Task<string?> TryComputeStreamSha256Async(Stream stream, CancellationToken cancellationToken) {
        if (!stream.CanSeek) return null;
        long position;
        try { position = stream.Position; } catch { return null; }
        try {
            stream.Position = 0;
            using IncrementalHash hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            var buffer = new byte[81920];
            int read;
            while ((read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) != 0)
                hash.AppendData(buffer, 0, read);
            cancellationToken.ThrowIfCancellationRequested();
            return ConvertToHexLower(hash.GetHashAndReset());
        } catch (OperationCanceledException) { throw; }
        catch { return null; }
        finally { try { stream.Position = position; } catch { } }
    }
}
