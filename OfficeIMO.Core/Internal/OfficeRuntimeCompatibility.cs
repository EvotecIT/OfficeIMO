using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Internal {
    /// <summary>Runtime operations shared by conversion orchestration on every supported framework.</summary>
    internal static partial class OfficeRuntimeCompatibility {
        internal static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
        internal static string TrimEndingDirectorySeparator(string path) {
            int rootLength = (Path.GetPathRoot(path) ?? string.Empty).Length;
            return path.Length > rootLength && (path[path.Length - 1] == Path.DirectorySeparatorChar || path[path.Length - 1] == Path.AltDirectorySeparatorChar)
                ? path.Substring(0, path.Length - 1) : path;
        }
#if NET8_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformGuard("windows")]
#endif
        internal static bool IsWindows => RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
#if NET8_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformGuard("linux")]
#endif
        internal static bool IsLinux => RuntimeInformation.IsOSPlatform(OSPlatform.Linux);
#if NET8_0_OR_GREATER
        [System.Runtime.Versioning.SupportedOSPlatformGuard("macos")]
#endif
        internal static bool IsMacOS => RuntimeInformation.IsOSPlatform(OSPlatform.OSX);
        internal static bool IsPathFullyQualified(string path) {
            if (!Path.IsPathRooted(path)) return false;
            if (!IsWindows) return true;
            string root = Path.GetPathRoot(path)!;
            return root.Length > 1 && !root.EndsWith(":", StringComparison.Ordinal);
        }

        internal const int MaximumArrayLength = 0X7FFFFFC7;

        internal static bool IsDefined<T>(T value) where T : struct, Enum {
#if NET8_0_OR_GREATER
            return Enum.IsDefined(value);
#else
            return Enum.IsDefined(typeof(T), value);
#endif
        }
        internal static bool IsAsciiLetterOrDigit(char value) => value >= 'a' && value <= 'z' || value >= 'A' && value <= 'Z' || value >= '0' && value <= '9';
        internal static bool IsAsciiHexDigit(char value) => value >= '0' && value <= '9' || value >= 'A' && value <= 'F' || value >= 'a' && value <= 'f';

        internal static byte[] HashSha256(byte[] bytes) {
#if NET8_0_OR_GREATER
            return SHA256.HashData(bytes);
#else
            using (var algorithm = SHA256.Create()) return algorithm.ComputeHash(bytes);
#endif
        }

        internal static byte[] HashSha256(byte[] bytes, int offset, int count) {
            using (var algorithm = SHA256.Create()) return algorithm.ComputeHash(bytes, offset, count);
        }

        internal static string ToHexString(byte[] bytes) {
#if NET8_0_OR_GREATER
            return Convert.ToHexString(bytes);
#else
            const string digits = "0123456789ABCDEF";
            var result = new char[checked(bytes.Length * 2)];
            for (int index = 0; index < bytes.Length; index++) {
                result[index * 2] = digits[bytes[index] >> 4];
                result[index * 2 + 1] = digits[bytes[index] & 15];
            }
            return new string(result);
#endif
        }

        internal static bool FixedTimeEquals(byte[] left, byte[] right) {
            if (left.Length != right.Length) return false;
            int difference = 0;
            for (int index = 0; index < left.Length; index++) difference |= left[index] ^ right[index];
            return difference == 0;
        }

        internal static string GetFullPath(string path, string baseDirectory) => Path.GetFullPath(Path.IsPathRooted(path) ? path : Path.Combine(baseDirectory, path));

        internal static string GetRelativePath(string baseDirectory, string path) {
#if NET8_0_OR_GREATER
            return Path.GetRelativePath(baseDirectory, path);
#else
            string root = Path.GetFullPath(baseDirectory);
            string target = Path.GetFullPath(path);
            StringComparison comparison = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            if (!string.Equals(Path.GetPathRoot(root), Path.GetPathRoot(target), comparison)) return target;
            char[] separators = { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar };
            string[] from = root.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            string[] to = target.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            int common = 0;
            while (common < from.Length && common < to.Length && string.Equals(from[common], to[common], comparison)) common++;
            string result = string.Join(Path.DirectorySeparatorChar.ToString(), Enumerable.Repeat("..", from.Length - common).Concat(to.Skip(common)));
            return result.Length == 0 ? "." : result;
#endif
        }

        internal static void WriteAll(Stream stream, byte[] bytes) => stream.Write(bytes, 0, bytes.Length);
        internal static Task WriteAllAsync(Stream stream, byte[] bytes, CancellationToken token) => stream.WriteAsync(bytes, 0, bytes.Length, token);

        internal static int ReadAtLeast(Stream stream, byte[] buffer, int minimum, bool throwOnEndOfStream = true) {
            int count = 0;
            while (count < minimum) {
                int read = stream.Read(buffer, count, buffer.Length - count);
                if (read == 0) {
                    if (throwOnEndOfStream) throw new EndOfStreamException();
                    break;
                }
                count += read;
            }
            return count;
        }

        /// <summary>Runs bounded workers without materializing the input selection; a failed worker cancels and joins its siblings.</summary>
        internal static async Task ForEachAsync<T>(IEnumerable<T> source, ParallelOptions options, Func<T, CancellationToken, Task> body) {
#if NET8_0_OR_GREATER
            await Parallel.ForEachAsync(source, options, async (item, token) => await body(item, token).ConfigureAwait(false)).ConfigureAwait(false);
#else
            using (var cancellation = CancellationTokenSource.CreateLinkedTokenSource(options.CancellationToken))
            using (var enumerator = source.GetEnumerator()) {
                var gate = new object();
                var workers = new Task[options.MaxDegreeOfParallelism];
                for (int index = 0; index < workers.Length; index++) {
                    workers[index] = Task.Run(async () => {
                        try {
                            while (true) {
                                T item;
                                lock (gate) {
                                    cancellation.Token.ThrowIfCancellationRequested();
                                    if (!enumerator.MoveNext()) return;
                                    item = enumerator.Current;
                                }
                                await body(item, cancellation.Token).ConfigureAwait(false);
                            }
                        } catch { cancellation.Cancel(); throw; }
                    }, CancellationToken.None);
                }
                await Task.WhenAll(workers).ConfigureAwait(false);
            }
#endif
        }
    }
}
