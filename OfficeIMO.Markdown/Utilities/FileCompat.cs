using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markdown;

/// <summary>
/// Small compatibility helpers for APIs missing on netstandard2.0.
/// </summary>
internal static class FileCompat {
#if NETSTANDARD2_0 || NET472 || NET48
    public static async Task WriteAllTextAsync(string path, string contents, Encoding encoding, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
        using var writer = new StreamWriter(stream, encoding);
        await writer.WriteAsync(contents).ConfigureAwait(false);
    }
#else
    public static Task WriteAllTextAsync(string path, string contents, Encoding encoding, CancellationToken cancellationToken) =>
        System.IO.File.WriteAllTextAsync(path, contents, encoding, cancellationToken);
#endif
}
