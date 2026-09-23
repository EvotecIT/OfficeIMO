using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Runtime.CompilerServices;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfActionPayloadFingerprint {
    private const int MaximumDepth = 32;
    private const int MaximumNodes = 4096;
    private static readonly ConditionalWeakTable<Dictionary<int, PdfIndirectObject>, PageNumberLookupCache> PageNumberLookups = new();
    private static readonly ConditionalWeakTable<Dictionary<int, PdfIndirectObject>, StreamHashCache> StreamHashes = new();
    private static readonly ConditionalWeakTable<Dictionary<int, PdfIndirectObject>, StringHashCache> StringHashes = new();
    private static readonly ConditionalWeakTable<Dictionary<int, PdfIndirectObject>, ReferenceHashCache> ReferenceHashes = new();

    internal static string? Create(
        PdfDictionary action,
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var builder = new StringBuilder();
        var activeReferences = new HashSet<(int ObjectNumber, int Generation)>();
        PageNumberLookup pageNumberLookup = PageNumberLookups.GetValue(
            objects,
            static source => new PageNumberLookupCache(source)).Get(limits, cancellationToken);
        if (!pageNumberLookup.IsComplete) return null;
        IReadOnlyDictionary<int, int> pageNumbers = pageNumberLookup.Value;
        int nodes = 0;
        bool complete = true;
        AppendDictionary(builder, action, objects, pageNumbers, activeReferences, depth: 0, ref nodes, ref complete, isActionRoot: true, useReferenceHashes: true, cancellationToken);
        return complete ? builder.ToString() : null;
    }

    private static void AppendObject(
        StringBuilder builder,
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyDictionary<int, int> pageNumbers,
        HashSet<(int ObjectNumber, int Generation)> activeReferences,
        int depth,
        ref int nodes,
        ref bool complete,
        bool useReferenceHashes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        nodes++;
        if (depth > MaximumDepth || nodes > MaximumNodes) {
            complete = false;
            return;
        }

        switch (value) {
            case null:
            case PdfNull:
                builder.Append('z');
                return;
            case PdfBoolean boolean:
                builder.Append(boolean.Value ? "b1" : "b0");
                return;
            case PdfNumber number:
                builder.Append('n').Append(number.Value.ToString("R", CultureInfo.InvariantCulture));
                return;
            case PdfName name:
                AppendText(builder, 'N', name.Name);
                return;
            case PdfStringObj text:
                AppendText(builder, 'S', StringHashes.GetValue(objects, static _ => new StringHashCache()).Get(text, cancellationToken));
                return;
            case PdfReference reference:
                AppendReference(builder, reference, objects, pageNumbers, activeReferences, depth, ref nodes, ref complete, useReferenceHashes, cancellationToken);
                return;
            case PdfArray array:
                builder.Append('[');
                for (int i = 0; i < array.Items.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AppendObject(builder, array.Items[i], objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete, useReferenceHashes, cancellationToken);
                    builder.Append(';');
                }
                builder.Append(']');
                return;
            case PdfDictionary dictionary:
                AppendDictionary(builder, dictionary, objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete, isActionRoot: false, useReferenceHashes, cancellationToken);
                return;
            case PdfStream stream:
                builder.Append("stream:");
                AppendDictionary(builder, stream.Dictionary, objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete, isActionRoot: false, useReferenceHashes, cancellationToken);
                AppendText(builder, 'H', StreamHashes.GetValue(objects, static _ => new StreamHashCache()).Get(stream, cancellationToken));
                return;
            default:
                AppendText(builder, '?', value.GetType().FullName ?? value.GetType().Name);
                return;
        }
    }

    private static void AppendReference(
        StringBuilder builder,
        PdfReference reference,
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyDictionary<int, int> pageNumbers,
        HashSet<(int ObjectNumber, int Generation)> activeReferences,
        int depth,
        ref int nodes,
        ref bool complete,
        bool useReferenceHashes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var key = (reference.ObjectNumber, reference.Generation);
        if (!activeReferences.Add(key)) {
            builder.Append("cycle:").Append(reference.ObjectNumber).Append(':').Append(reference.Generation);
            return;
        }
        try {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                builder.Append("ref:").Append(reference.ObjectNumber).Append(':').Append(reference.Generation);
                return;
            }
            if (indirect.Value is PdfDictionary dictionary &&
                string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal)) {
                if (pageNumbers.TryGetValue(reference.ObjectNumber, out int pageNumber)) {
                    builder.Append("page:").Append(pageNumber);
                } else {
                    builder.Append("page-ref:").Append(reference.ObjectNumber).Append(':').Append(reference.Generation);
                }
                return;
            }
            if (useReferenceHashes) {
                ReferenceHashResult result = ReferenceHashes.GetValue(objects, static _ => new ReferenceHashCache()).Get(
                    reference,
                    depth + 1,
                    () => CreateReferenceHash(indirect.Value, objects, pageNumbers, key, depth + 1, cancellationToken), cancellationToken);
                nodes = checked(nodes + result.Nodes);
                if (!result.Complete || nodes > MaximumNodes) {
                    complete = false;
                    return;
                }
                AppendText(builder, 'R', result.Hash);
                return;
            }
            AppendObject(builder, indirect.Value, objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete, useReferenceHashes: false, cancellationToken);
        } finally {
            activeReferences.Remove(key);
        }
    }

    private static void AppendDictionary(
        StringBuilder builder,
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyDictionary<int, int> pageNumbers,
        HashSet<(int ObjectNumber, int Generation)> activeReferences,
        int depth,
        ref int nodes,
        ref bool complete,
        bool isActionRoot,
        bool useReferenceHashes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (dictionary.Items.Count > MaximumNodes + (isActionRoot ? 2 : 0)) {
            complete = false;
            return;
        }
        builder.Append('{');
        var keys = new List<string>(dictionary.Items.Count);
        foreach (string key in dictionary.Items.Keys) {
            cancellationToken.ThrowIfCancellationRequested();
            keys.Add(key);
        }
        try {
            keys.Sort((left, right) => {
                cancellationToken.ThrowIfCancellationRequested();
                return StringComparer.Ordinal.Compare(left, right);
            });
        } catch (InvalidOperationException error) when (error.InnerException is OperationCanceledException) {
            cancellationToken.ThrowIfCancellationRequested();
            throw;
        }
        cancellationToken.ThrowIfCancellationRequested();
        foreach (string key in keys) {
            cancellationToken.ThrowIfCancellationRequested();
            if (isActionRoot &&
                (string.Equals(key, "S", StringComparison.Ordinal) ||
                 string.Equals(key, "Next", StringComparison.Ordinal))) continue;
            AppendText(builder, 'K', key);
            AppendObject(builder, dictionary.Items[key], objects, pageNumbers, activeReferences, depth + 1, ref nodes, ref complete, useReferenceHashes, cancellationToken);
            builder.Append(';');
        }
        builder.Append('}');
    }

    private static ReferenceHashResult CreateReferenceHash(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyDictionary<int, int> pageNumbers,
        (int ObjectNumber, int Generation) rootReference,
        int depth,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var builder = new StringBuilder();
        var activeReferences = new HashSet<(int ObjectNumber, int Generation)> { rootReference };
        int nodes = 0;
        bool complete = true;
        AppendObject(builder, value, objects, pageNumbers, activeReferences, depth, ref nodes, ref complete, useReferenceHashes: false, cancellationToken);
        if (!complete) return new ReferenceHashResult(string.Empty, nodes, complete: false);
        string hash = HashUtf8(builder, cancellationToken);
        return new ReferenceHashResult(hash, nodes, complete: true);
    }

    private static PageNumberLookup BuildPageNumberLookup(Dictionary<int, PdfIndirectObject> objects, PdfReadLimits limits, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var result = new Dictionary<int, int>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, cancellationToken: cancellationToken);
        if (catalog == null || !catalog.Items.TryGetValue("Pages", out PdfObject? pages)) return new PageNumberLookup(result, isComplete: true);
        var visited = new HashSet<int>();
        bool complete = true;
        AddPageTreeNode(pages, objects, visited, result, depth: 0, limits, ref complete, cancellationToken);
        return new PageNumberLookup(result, complete);
    }

    private static void AddPageTreeNode(
        PdfObject node,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> visited,
        Dictionary<int, int> pageNumbers,
        int depth,
        PdfReadLimits limits,
        ref bool complete,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (depth > limits.MaxPageTreeDepth) {
            complete = false;
            return;
        }
        int objectNumber = 0;
        if (node is PdfReference reference) {
            objectNumber = reference.ObjectNumber;
            if (!visited.Add(objectNumber) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) return;
            if (visited.Count > limits.MaxPageTreeNodes) {
                complete = false;
                return;
            }
            node = indirect.Value;
        }
        if (node is not PdfDictionary dictionary) return;
        if (string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal)) {
            if (objectNumber > 0 && !pageNumbers.ContainsKey(objectNumber)) pageNumbers.Add(objectNumber, pageNumbers.Count + 1);
            return;
        }
        if (!dictionary.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            PdfObjectLookup.Resolve(objects, kidsObject) is not PdfArray kids) return;
        for (int index = 0; index < kids.Items.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            AddPageTreeNode(kids.Items[index], objects, visited, pageNumbers, depth + 1, limits, ref complete, cancellationToken);
        }
    }

    private static void AppendText(StringBuilder builder, char prefix, string value) =>
        builder.Append(prefix).Append(value.Length).Append(':').Append(value);

    private static string HashBase64(byte[] bytes, CancellationToken cancellationToken) {
        if (!cancellationToken.CanBeCanceled) {
#if NET8_0_OR_GREATER
            return Convert.ToBase64String(SHA256.HashData(bytes));
#else
            using var sha256 = SHA256.Create();
            return Convert.ToBase64String(sha256.ComputeHash(bytes));
#endif
        }

        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        const int chunkSize = 64 * 1024;
        for (int offset = 0; offset < bytes.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(chunkSize, bytes.Length - offset);
            hash.AppendData(bytes, offset, count);
            offset += count;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return Convert.ToBase64String(hash.GetHashAndReset());
    }

    private static string HashUtf8(StringBuilder builder, CancellationToken cancellationToken) {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var characters = new char[16 * 1024];
        var bytes = new byte[Encoding.UTF8.GetMaxByteCount(characters.Length)];
        var encoder = Encoding.UTF8.GetEncoder();
        for (int offset = 0; offset < builder.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(characters.Length, builder.Length - offset);
            builder.CopyTo(offset, characters, 0, count);
            int written = encoder.GetBytes(characters, 0, count, bytes, 0, offset + count == builder.Length);
            hash.AppendData(bytes, 0, written);
            offset += count;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return Convert.ToBase64String(hash.GetHashAndReset());
    }

    private sealed class PageNumberLookup {
        internal PageNumberLookup(IReadOnlyDictionary<int, int> value, bool isComplete) { Value = value; IsComplete = isComplete; }
        internal IReadOnlyDictionary<int, int> Value { get; }
        internal bool IsComplete { get; }
    }

    private sealed class PageNumberLookupCache {
        private readonly Dictionary<int, PdfIndirectObject> _objects;
        private readonly Dictionary<(int Depth, int Nodes), PdfReadCache<PageNumberLookup>> _values = new();

        internal PageNumberLookupCache(Dictionary<int, PdfIndirectObject> objects) { _objects = objects; }

        internal PageNumberLookup Get(PdfReadLimits limits, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            var key = (limits.MaxPageTreeDepth, limits.MaxPageTreeNodes);
            PdfReadCache<PageNumberLookup> entry;
            lock (_values) {
                if (!_values.TryGetValue(key, out entry!)) {
                    entry = new PdfReadCache<PageNumberLookup>();
                    _values.Add(key, entry);
                }
            }
            return entry.GetOrCreate((Objects: _objects, Limits: limits),
                static (state, token) => BuildPageNumberLookup(state.Objects, state.Limits, token), cancellationToken);
        }
    }

    private sealed class StreamHashCache {
        private readonly Dictionary<PdfStream, PdfReadCache<string>> _values = new();

        internal string Get(PdfStream stream, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfReadCache<string> entry;
            lock (_values) {
                if (!_values.TryGetValue(stream, out entry!)) {
                    entry = new PdfReadCache<string>();
                    _values.Add(stream, entry);
                }
            }
            return entry.GetOrCreate(stream, static (source, token) => HashBase64(source.GetData(token), token), cancellationToken);
        }
    }

    private sealed class StringHashCache {
        private readonly Dictionary<PdfStringObj, PdfReadCache<string>> _values = new();

        internal string Get(PdfStringObj text, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfReadCache<string> entry;
            lock (_values) {
                if (!_values.TryGetValue(text, out entry!)) {
                    entry = new PdfReadCache<string>();
                    _values.Add(text, entry);
                }
            }
            return entry.GetOrCreate(text, static (source, token) => HashBase64(source.RawBytes, token), cancellationToken);
        }
    }

    private readonly struct ReferenceHashResult {
        internal ReferenceHashResult(string hash, int nodes, bool complete) { Hash = hash; Nodes = nodes; Complete = complete; }
        internal string Hash { get; }
        internal int Nodes { get; }
        internal bool Complete { get; }
    }

    private sealed class ReferenceHashCache {
        private readonly Dictionary<(int ObjectNumber, int Generation, int Depth), ReferenceHashResult> _values = new();

        internal ReferenceHashResult Get(PdfReference reference, int depth, Func<ReferenceHashResult> create, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            var key = (reference.ObjectNumber, reference.Generation, depth);
            lock (_values) {
                if (_values.TryGetValue(key, out ReferenceHashResult value)) return value;
            }
            ReferenceHashResult created = create();
            lock (_values) {
                if (_values.TryGetValue(key, out ReferenceHashResult value)) return value;
                _values.Add(key, created);
            }
            return created;
        }
    }
}
