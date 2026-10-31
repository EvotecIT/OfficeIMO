using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Internal {
#if NETSTANDARD2_0 || NETFRAMEWORK
    /// <summary>Older-framework overloads used by bounded conversion orchestration.</summary>
    internal static class OfficeCollectionCompatibility {
        internal static IEnumerable<T> DistinctBy<T, TKey>(this IEnumerable<T> source, Func<T, TKey> selector) {
            var keys = new HashSet<TKey>();
            foreach (T item in source) if (keys.Add(selector(item))) yield return item;
        }
        internal static IEnumerable<T> Order<T>(this IEnumerable<T> source, IComparer<T> comparer) => source.OrderBy(item => item, comparer);
        internal static TValue? GetValueOrDefault<TKey, TValue>(this IReadOnlyDictionary<TKey, TValue> source, TKey key) => source.TryGetValue(key, out TValue value) ? value : default;
        internal static Task DisposeAsync(this Stream stream) { stream.Dispose(); return Task.CompletedTask; }
        internal static async Task ReadExactlyAsync(this Stream stream, byte[] bytes, CancellationToken token) {
            int count = 0;
            while (count < bytes.Length) {
                int read = await stream.ReadAsync(bytes, count, bytes.Length - count, token).ConfigureAwait(false);
                if (read == 0) throw new EndOfStreamException();
                count += read;
            }
        }
        internal static async Task<T> WaitAsync<T>(this Task<T> task, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            var cancelled = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            using (token.Register(() => cancelled.TrySetResult(true))) {
                if (await Task.WhenAny(task, cancelled.Task).ConfigureAwait(false) != task) token.ThrowIfCancellationRequested();
                return await task.ConfigureAwait(false);
            }
        }
#if NETSTANDARD2_0
        internal static HashSet<T> ToHashSet<T>(this IEnumerable<T> source, IEqualityComparer<T>? comparer = null) => new HashSet<T>(source, comparer);
#endif
        internal static bool TryAdd<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TKey key, TValue value) {
            if (dictionary.ContainsKey(key)) return false;
            dictionary.Add(key, value);
            return true;
        }
        internal static Task<int> ReadAsync(this Stream stream, byte[] buffer, CancellationToken token) => stream.ReadAsync(buffer, 0, buffer.Length, token);
        internal static Task WriteAsync(this Stream stream, byte[] buffer, CancellationToken token) => stream.WriteAsync(buffer, 0, buffer.Length, token);
        internal static Task CopyToAsync(this Stream source, Stream destination, CancellationToken token) => source.CopyToAsync(destination, 81920, token);
        internal static async Task<string> ReadToEndAsync(this TextReader reader, CancellationToken token) {
            var result = new StringBuilder();
            var buffer = new char[4096];
            int count;
            token.ThrowIfCancellationRequested();
            while ((count = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) != 0) {
                token.ThrowIfCancellationRequested();
                result.Append(buffer, 0, count);
            }
            token.ThrowIfCancellationRequested();
            return result.ToString();
        }
        internal static void Write(this Stream stream, byte[] bytes) => stream.Write(bytes, 0, bytes.Length);
        internal static void ReadExactly(this Stream stream, byte[] buffer) => OfficeRuntimeCompatibility.ReadAtLeast(stream, buffer, buffer.Length);
        internal static int ReadAtLeast(this Stream stream, byte[] buffer, int minimumBytes, bool throwOnEndOfStream = true) => OfficeRuntimeCompatibility.ReadAtLeast(stream, buffer, minimumBytes, throwOnEndOfStream);
    }
#endif
}
