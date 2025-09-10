using System.Text;
using System.Threading.Tasks;

namespace OfficeIMO.Markdown;

/// <summary>
/// Small compatibility helpers for APIs missing on netstandard2.0.
/// </summary>
internal static class FileCompat {
#if NETSTANDARD2_0 || NET472 || NET48
    public static Task WriteAllTextAsync(string path, string contents, Encoding encoding) {
        System.IO.File.WriteAllText(path, contents, encoding);
        return Task.CompletedTask;
    }
#else
    public static Task WriteAllTextAsync(string path, string contents, Encoding encoding) => System.IO.File.WriteAllTextAsync(path, contents, encoding);
#endif
}

