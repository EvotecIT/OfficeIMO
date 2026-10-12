using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

/// <summary>
/// Keeps decoded and prefiltered source images between renders of the same drawing, such as the tiles and zoom levels
/// of one page. Entries are keyed by the identity of their source (an image's encoded bytes, or a decoded image and
/// its filtered size), so the cache belongs to drawings that are not modified while it is in use. The least recently
/// used entries are released once the decoded pixels exceed the capacity. Cached images are shared and never modified.
/// </summary>
internal sealed class OfficeRasterImageCache {
    private readonly object _sync = new();
    private readonly long _capacityBytes;
    private readonly Dictionary<Key, LinkedListNode<Entry>> _entries = new();
    private readonly LinkedList<Entry> _recent = new();
    private long _usedBytes;

    internal OfficeRasterImageCache(long capacityBytes) {
        if (capacityBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(capacityBytes), "The cache capacity must be positive.");
        _capacityBytes = capacityBytes;
    }

    internal long UsedBytes {
        get { lock (_sync) return _usedBytes; }
    }

    /// <summary>A decoded image for its encoded bytes.</summary>
    internal bool TryGetDecoded(byte[] encoded, out OfficeRasterImage image) => TryGet(new Key(encoded, 0, 0), out image);

    internal void AddDecoded(byte[] encoded, OfficeRasterImage image) => Add(new Key(encoded, 0, 0), image);

    /// <summary>A decoded image reduced to a sampling size.</summary>
    internal bool TryGetFiltered(OfficeRasterImage source, int width, int height, out OfficeRasterImage image) =>
        TryGet(new Key(source, width, height), out image);

    internal void AddFiltered(OfficeRasterImage source, int width, int height, OfficeRasterImage image) =>
        Add(new Key(source, width, height), image);

    private bool TryGet(Key key, out OfficeRasterImage image) {
        lock (_sync) {
            if (_entries.TryGetValue(key, out LinkedListNode<Entry>? node)) {
                _recent.Remove(node);
                _recent.AddFirst(node);
                image = node.Value.Image;
                return true;
            }
        }
        image = null!;
        return false;
    }

    private void Add(Key key, OfficeRasterImage image) {
        long bytes = (long)image.Width * image.Height * 4L;
        if (bytes > _capacityBytes) return;
        lock (_sync) {
            if (_entries.ContainsKey(key)) return;
            while (_usedBytes + bytes > _capacityBytes && _recent.Last is { } oldest) {
                _recent.RemoveLast();
                _entries.Remove(oldest.Value.Key);
                _usedBytes -= oldest.Value.Bytes;
            }
            _entries[key] = _recent.AddFirst(new Entry(key, image, bytes));
            _usedBytes += bytes;
        }
    }

    private readonly struct Key : IEquatable<Key> {
        private readonly object _source;
        private readonly int _width;
        private readonly int _height;

        internal Key(object source, int width, int height) {
            _source = source;
            _width = width;
            _height = height;
        }

        public bool Equals(Key other) => ReferenceEquals(_source, other._source) && _width == other._width && _height == other._height;

        public override bool Equals(object? obj) => obj is Key other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                return (RuntimeHelpers.GetHashCode(_source) * 397 ^ _width) * 397 ^ _height;
            }
        }
    }

    private sealed class Entry {
        internal Entry(Key key, OfficeRasterImage image, long bytes) {
            Key = key;
            Image = image;
            Bytes = bytes;
        }

        internal Key Key { get; }
        internal OfficeRasterImage Image { get; }
        internal long Bytes { get; }
    }
}
