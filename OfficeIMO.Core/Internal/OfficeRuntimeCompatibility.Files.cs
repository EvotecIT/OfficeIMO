using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Internal {
    internal static partial class OfficeRuntimeCompatibility {
        internal static async Task<byte[]> ReadAllBytesAsync(string path, CancellationToken token) {
#if NET8_0_OR_GREATER
            return await File.ReadAllBytesAsync(path, token).ConfigureAwait(false);
#else
            using (var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.Asynchronous))
            using (var result = new MemoryStream()) {
                await source.CopyToAsync(result, 81920, token).ConfigureAwait(false);
                return result.ToArray();
            }
#endif
        }
        internal static async Task WriteAllBytesAsync(string path, byte[] bytes, CancellationToken token) {
#if NET8_0_OR_GREATER
            await File.WriteAllBytesAsync(path, bytes, token).ConfigureAwait(false);
#else
            token.ThrowIfCancellationRequested();
            using (var target = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read, 81920, FileOptions.Asynchronous))
                await target.WriteAsync(bytes, 0, bytes.Length, token).ConfigureAwait(false);
#endif
        }
    }
}

namespace OfficeIMO.Internal {
    internal static partial class OfficeRuntimeCompatibility {
        internal static System.Collections.Generic.IEnumerable<string> EnumerateFiles(string path, string pattern, bool recursive) {
#if NET8_0_OR_GREATER
            return Directory.EnumerateFiles(path, pattern, new EnumerationOptions {
                RecurseSubdirectories = recursive, IgnoreInaccessible = false,
                AttributesToSkip = FileAttributes.ReparsePoint, ReturnSpecialDirectories = false
            });
#else
            return EnumerateWithoutLinks(path, pattern, recursive);
#endif
        }
#if !NET8_0_OR_GREATER
        private static System.Collections.Generic.IEnumerable<string> EnumerateWithoutLinks(string path, string pattern, bool recursive) {
            var pending = new System.Collections.Generic.Stack<string>();
            pending.Push(path);
            while (pending.Count != 0) {
                string directory = pending.Pop();
                if (directory != path && (File.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0) continue;
                foreach (string file in Directory.EnumerateFiles(directory, pattern, SearchOption.TopDirectoryOnly))
                    if ((File.GetAttributes(file) & FileAttributes.ReparsePoint) == 0) yield return file;
                if (recursive)
                    foreach (string child in Directory.EnumerateDirectories(directory))
                        if ((File.GetAttributes(child) & FileAttributes.ReparsePoint) == 0) pending.Push(child);
            }
        }
#endif
    }
}
