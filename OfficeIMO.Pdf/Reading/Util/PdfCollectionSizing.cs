#if NET8_0_OR_GREATER
using System.Buffers;
using System.Runtime.CompilerServices;
#endif

namespace OfficeIMO.Pdf;

internal static class PdfCollectionSizing {
    internal static int BoundedInitialCapacity(
        int expectedCount,
        int maximum) {
        if (expectedCount <= 0 || maximum <= 0) return 0;
        return Math.Min(expectedCount, maximum);
    }

    internal static HashSet<T> CreateHashSet<T>(
        int expectedCount,
        int maximum) {
#if NETSTANDARD2_0 || NET472
        return new HashSet<T>();
#else
        return new HashSet<T>(BoundedInitialCapacity(expectedCount, maximum));
#endif
    }
}

/// <summary>
/// Grows transient value collections through the shared pool, then materializes one
/// exactly-sized managed array for the retained parse result.
/// </summary>
internal sealed class PdfPooledValueBuilder<T> : IDisposable where T : struct {
#if NET8_0_OR_GREATER
    private static readonly bool ClearOnReturn = RuntimeHelpers.IsReferenceOrContainsReferences<T>();
    private T[]? _buffer;
    private int _count;
#else
    private readonly List<T> _items;
#endif

    internal PdfPooledValueBuilder(int initialCapacity) {
#if NET8_0_OR_GREATER
        _buffer = ArrayPool<T>.Shared.Rent(Math.Max(1, initialCapacity));
#else
        _items = new List<T>(initialCapacity);
#endif
    }

    internal int Count {
        get {
#if NET8_0_OR_GREATER
            return _count;
#else
            return _items.Count;
#endif
        }
    }

    internal void Add(T value) {
#if NET8_0_OR_GREATER
        T[] buffer = _buffer ?? throw new ObjectDisposedException(nameof(PdfPooledValueBuilder<T>));
        if (_count == buffer.Length) {
            T[] expanded = ArrayPool<T>.Shared.Rent(checked(buffer.Length * 2));
            Array.Copy(buffer, expanded, _count);
            ArrayPool<T>.Shared.Return(buffer, clearArray: ClearOnReturn);
            _buffer = buffer = expanded;
        }

        buffer[_count++] = value;
#else
        _items.Add(value);
#endif
    }

    internal T[] ToArray() {
#if NET8_0_OR_GREATER
        if (_count == 0) return Array.Empty<T>();
        var result = new T[_count];
        Array.Copy(_buffer!, result, _count);
        return result;
#else
        return _items.ToArray();
#endif
    }

    public void Dispose() {
#if NET8_0_OR_GREATER
        if (_buffer is not null) {
            ArrayPool<T>.Shared.Return(_buffer, clearArray: ClearOnReturn);
            _buffer = null;
            _count = 0;
        }
#endif
    }
}
