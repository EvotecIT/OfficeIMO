using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Bounds-checks the self-contained AV1 item path before any entropy decoder allocates image planes.</summary>
internal sealed partial class OfficeAvifContainerReader {
    private const int MaximumBoxes = 4096;
    private const int MaximumItems = 1024;
    private readonly byte[] _bytes;
    private readonly OfficeRasterDecodeOptions _options;
    private readonly Dictionary<uint, string> _types = new();
    private readonly Dictionary<uint, (int Offset, int Length)> _locations = new();
    private readonly Dictionary<uint, List<(int Index, bool Essential)>> _associations = new();
    private readonly List<Box> _properties = new();
    private readonly List<Box> _dataBoxes = new();
    private readonly List<(uint From, uint To)> _alphaReferences = new();
    private readonly HashSet<uint> _premultipliedItems = new();
    private uint _primaryId;
    private int _boxCount;

    private OfficeAvifContainerReader(byte[] bytes, OfficeRasterDecodeOptions options) {
        _bytes = bytes;
        _options = options;
    }

    /// <summary>Reads whole, untransformed Main 8/10-bit color/alpha items. Container support is distinct from pixels.</summary>
    internal static bool TryRead(byte[]? bytes, OfficeRasterDecodeOptions options, out OfficeAvifContainer? container) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate();
        options.CancellationToken.ThrowIfCancellationRequested();
        container = null;
        if (bytes == null || bytes.Length < 16 || bytes.Length > options.MaximumEncodedBytes) return false;
        try {
            var reader = new OfficeAvifContainerReader(bytes, options);
            container = reader.Read();
            return true;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    private OfficeAvifContainer Read() {
        Box? meta = null;
        bool brand = false;
        bool hasFileType = false;
        foreach (Box box in Boxes(0, _bytes.Length)) {
            if (box.Type == "ftyp") {
                Require(!hasFileType && box.Length >= 8 && box.Length % 4 == 0);
                hasFileType = true;
                brand = Type(box.Start) == "avif";
                for (int p = box.Start + 8; p < box.End; p += 4) brand |= Type(p) == "avif";
            } else if (box.Type == "meta") {
                Require(meta == null);
                meta = box;
            } else if (box.Type == "mdat") {
                _dataBoxes.Add(box);
            } else if (box.Type == "moov") {
                throw new FormatException("Image sequences are outside the still-item path.");
            }
        }
        Require(hasFileType && brand && meta != null && _dataBoxes.Count > 0);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        int children = FullBox(meta!.Value, 0, out uint flags);
        Require(flags == 0);
        foreach (Box box in Boxes(children, meta.Value.End)) {
            if (box.Type is "pitm" or "iloc" or "iinf" or "iprp" or "iref") Require(seen.Add(box.Type));
            switch (box.Type) {
                case "pitm": ReadPrimary(box); break;
                case "iloc": ReadLocations(box); break;
                case "iinf": ReadItemTypes(box); break;
                case "iprp": ReadProperties(box); break;
                case "iref": ReadReferences(box); break;
            }
        }
        Require(seen.Contains("pitm") && seen.Contains("iloc") && seen.Contains("iinf") && seen.Contains("iprp"));
        Require(!_premultipliedItems.Contains(_primaryId));
        OfficeAvifImageItem color = ResolveItem(_primaryId, alpha: false);
        OfficeAvifImageItem? alpha = null;
        foreach (var reference in _alphaReferences) {
            if (reference.To != _primaryId) continue;
            Require(reference.From != _primaryId && alpha == null);
            alpha = ResolveItem(reference.From, alpha: true);
            Require(alpha.Width == color.Width && alpha.Height == color.Height && alpha.BitDepth == color.BitDepth);
        }
        return new OfficeAvifContainer(color, alpha);
    }

    private IEnumerable<Box> Boxes(int start, int end) {
        for (int p = start; p < end;) {
            _options.CancellationToken.ThrowIfCancellationRequested();
            Require(++_boxCount <= MaximumBoxes && end - p >= 8);
            ulong length = U32(p);
            string type = Type(p + 4);
            int header = 8;
            if (length == 1) {
                Require(end - p >= 16);
                length = U64(p + 8);
                header = 16;
            } else if (length == 0) {
                length = (ulong)(end - p);
            }
            Require(length >= (ulong)header && length <= (ulong)(end - p));
            int next = checked(p + (int)length);
            yield return new Box(type, p + header, next);
            p = next;
        }
    }

    private int FullBox(Box box, byte maximumVersion, out uint flags) {
        Require(box.Length >= 4 && _bytes[box.Start] <= maximumVersion);
        flags = U32(box.Start) & 0xFFFFFFU;
        return box.Start + 4;
    }

    private uint U32(int p) {
        Require(p >= 0 && p <= _bytes.Length - 4);
        return (uint)_bytes[p] << 24 | (uint)_bytes[p + 1] << 16 | (uint)_bytes[p + 2] << 8 | _bytes[p + 3];
    }

    private ulong U64(int p) => (ulong)U32(p) << 32 | U32(p + 4);

    private string Type(int p) {
        Require(p >= 0 && p <= _bytes.Length - 4);
        return System.Text.Encoding.ASCII.GetString(_bytes, p, 4);
    }

    private static void Require(bool condition) {
        if (!condition) throw new FormatException("Invalid or unsupported AVIF item container.");
    }

    private readonly struct Box {
        internal Box(string type, int start, int end) { Type = type; Start = start; End = end; }
        internal string Type { get; }
        internal int Start { get; }
        internal int End { get; }
        internal int Length => End - Start;
    }

    /// <summary>Reads a box-local integer without allowing its cursor into an adjacent sibling.</summary>
    private sealed class Cursor {
        private readonly OfficeAvifContainerReader _owner;
        private readonly int _end;
        internal Cursor(OfficeAvifContainerReader owner, int start, int end) { _owner = owner; Position = start; _end = end; }
        internal int Position { get; private set; }
        internal int Remaining => _end - Position;
        internal ulong Integer(int size) {
            Require(size >= 0 && size <= 8 && Remaining >= size);
            ulong value = 0;
            for (int i = 0; i < size; i++) value = value << 8 | _owner._bytes[Position++];
            return value;
        }
        internal uint Id(bool large) => checked((uint)Integer(large ? 4 : 2));
        internal string Type() {
            Require(Remaining >= 4);
            string value = _owner.Type(Position);
            Position += 4;
            return value;
        }
        internal string TerminatedString() {
            int start = Position;
            SkipTerminatedString();
            Require(Position - start <= 4096);
            string value = System.Text.Encoding.UTF8.GetString(_owner._bytes, start, Position - start - 1);
            return value;
        }
        internal void SkipTerminatedString() {
            while (Position < _end && _owner._bytes[Position] != 0) {
                if ((Position & 4095) == 0) _owner._options.CancellationToken.ThrowIfCancellationRequested();
                Position++;
            }
            Require(Position < _end);
            Position++;
        }
        internal void End() => Require(Remaining == 0);
    }
}
