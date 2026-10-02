using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextHyphenationPatterns {
    // Only fixed embedded resources are parsed. Lookup state is immutable after Lazy publication;
    // token data and computed weights are local, so calls neither grow a cache nor share buffers.
    private sealed class PatternDictionary {
        private readonly Dictionary<long, int> _edges = new Dictionary<long, int>();
        private readonly List<byte[]?> _weights = new List<byte[]?> { null };
        private readonly Dictionary<string, int[]> _exceptions = new Dictionary<string, int[]>(StringComparer.Ordinal);
        internal int LeftMinimum { get; }
        internal int RightMinimum { get; }

        private PatternDictionary(int left, int right) { LeftMinimum = left; RightMinimum = right; }

        internal static PatternDictionary Load(string resource, int left, int right) {
            var dictionary = new PatternDictionary(left, right);
            using (Stream stream = typeof(OfficeTextHyphenationPatterns).Assembly.GetManifestResourceStream(
                "OfficeIMO.Core.Hyphenation." + resource) ?? throw new InvalidOperationException("Missing embedded hyphenation resource."))
            using (var reader = new StreamReader(stream, Encoding.UTF8)) {
                string? line;
                while ((line = reader.ReadLine()) != null) {
                    if (line.Length == 0 || line[0] == '#') continue;
                    if (line[0] == '!') dictionary.AddException(line.Substring(1));
                    else dictionary.AddPattern(line);
                }
            }
            return dictionary;
        }

        private void AddException(string markedWord) {
            var word = new StringBuilder();
            var points = new List<int>();
            foreach (char value in markedWord) {
                if (value == '-') points.Add(word.Length);
                else word.Append(value);
            }
            _exceptions.Add(word.ToString(), points.ToArray());
        }

        private void AddPattern(string pattern) {
            int node = 0;
            var weights = new List<byte> { 0 };
            foreach (char value in pattern) {
                if (value >= '0' && value <= '9') { weights[weights.Count - 1] = (byte)(value - '0'); continue; }
                long key = ((long)node << 16) | value;
                if (!_edges.TryGetValue(key, out int next)) {
                    next = _weights.Count;
                    _edges.Add(key, next);
                    _weights.Add(null);
                }
                node = next;
                weights.Add(0);
            }
            byte[]? existing = _weights[node];
            if (existing == null) _weights[node] = weights.ToArray();
            else for (int index = 0; index < weights.Count; index++) existing[index] = Math.Max(existing[index], weights[index]);
        }

        internal IReadOnlyList<int> GetBreakpoints(string word) {
            if (_exceptions.TryGetValue(word, out int[]? exception)) return exception;
            string padded = "." + word + ".";
            var values = new byte[padded.Length + 1];
            for (int start = 0; start < padded.Length; start++) {
                int node = 0;
                for (int end = start; end < padded.Length; end++) {
                    if (!_edges.TryGetValue(((long)node << 16) | padded[end], out node)) break;
                    byte[]? weights = _weights[node];
                    if (weights == null) continue;
                    for (int index = 0; index < weights.Length; index++)
                        values[start + index] = Math.Max(values[start + index], weights[index]);
                }
            }
            var result = new List<int>();
            for (int point = 1; point < word.Length; point++) if ((values[point + 1] & 1) != 0) result.Add(point);
            return result;
        }
    }
}
