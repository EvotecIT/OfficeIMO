using System;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Shared HVAR advance-width evaluator for TrueType and CFF2 variable fonts.</summary>
internal sealed class OfficeOpenTypeHvarMetrics {
    private readonly OfficeOpenTypeItemVariationStore _store;
    private readonly DeltaSetIndexMap? _advanceMap;

    private OfficeOpenTypeHvarMetrics(
        OfficeOpenTypeItemVariationStore store,
        DeltaSetIndexMap? advanceMap) {
        _store = store;
        _advanceMap = advanceMap;
    }

    internal static OfficeOpenTypeHvarMetrics? TryParse(
        OfficeOpenTypeReader reader,
        OfficeFontVariationModel model) {
        if (!reader.TryGetTable("HVAR", out int offset, out int length)) return null;
        int end = checked(offset + length);
        if (length < 20 || reader.ReadUInt16(offset) != 1 || reader.ReadUInt16(offset + 2) != 0) {
            throw new InvalidDataException("The HVAR table header is invalid.");
        }
        uint storeRelative = reader.ReadUInt32(offset + 4);
        uint advanceMapRelative = reader.ReadUInt32(offset + 8);
        if (storeRelative == 0 || storeRelative > int.MaxValue || advanceMapRelative > int.MaxValue) {
            throw new InvalidDataException("The HVAR table offsets are invalid.");
        }
        int storeOffset = checked(offset + (int)storeRelative);
        if (storeOffset < offset || storeOffset >= end) {
            throw new InvalidDataException("The HVAR ItemVariationStore offset is invalid.");
        }
        OfficeOpenTypeItemVariationStore store = OfficeOpenTypeItemVariationStore.Parse(
            reader,
            storeOffset,
            end,
            model);
        DeltaSetIndexMap? map = advanceMapRelative == 0
            ? null
            : DeltaSetIndexMap.Parse(reader, checked(offset + (int)advanceMapRelative), end);
        if (map == null) {
            // OpenType maps an absent advanceWidthMapping to outer index zero and
            // inner index glyphId. Validate that complete implicit range while the
            // font is being registered instead of deferring a malformed index to
            // the first measurement or outline request.
            store.ValidateIndex(0, reader.GlyphCount - 1);
        } else {
            map.Validate(store);
        }
        return new OfficeOpenTypeHvarMetrics(store, map);
    }

    internal int AdvanceWidthDelta(int glyphId) {
        DeltaSetIndex index = _advanceMap?.Resolve(glyphId) ?? new DeltaSetIndex(0, glyphId);
        return _store.Evaluate(index.Outer, index.Inner);
    }

    private sealed class DeltaSetIndexMap {
        private readonly DeltaSetIndex[] _entries;
        private DeltaSetIndexMap(DeltaSetIndex[] entries) => _entries = entries;

        internal static DeltaSetIndexMap Parse(OfficeOpenTypeReader reader, int offset, int end) {
            if (offset < 0 || offset > end - 4) {
                throw new InvalidDataException("An HVAR DeltaSetIndexMap is truncated.");
            }
            int format = reader.Data[offset];
            int entryFormat = reader.Data[offset + 1];
            int entrySize = ((entryFormat >> 4) & 0x03) + 1;
            int innerBits = (entryFormat & 0x0F) + 1;
            int count;
            int cursor;
            if (format == 0) {
                count = reader.ReadUInt16(offset + 2);
                cursor = offset + 4;
            } else if (format == 1) {
                if (offset > end - 6) throw new InvalidDataException("An HVAR DeltaSetIndexMap is truncated.");
                uint countValue = reader.ReadUInt32(offset + 2);
                if (countValue > 1_000_000) throw new InvalidDataException("An HVAR DeltaSetIndexMap is too large.");
                count = (int)countValue;
                cursor = offset + 6;
            } else {
                throw new NotSupportedException("The HVAR DeltaSetIndexMap format is not supported.");
            }
            if (count <= 0 || count > 1_000_000 || cursor > end - checked(count * entrySize)) {
                throw new InvalidDataException("An HVAR DeltaSetIndexMap directory is invalid.");
            }
            uint innerMask = (1U << innerBits) - 1U;
            var entries = new DeltaSetIndex[count];
            for (int item = 0; item < count; item++) {
                uint value = 0;
                for (int index = 0; index < entrySize; index++) value = (value << 8) | reader.Data[cursor++];
                entries[item] = new DeltaSetIndex((int)(value >> innerBits), (int)(value & innerMask));
            }
            return new DeltaSetIndexMap(entries);
        }

        internal DeltaSetIndex Resolve(int glyphId) =>
            _entries[Math.Min(Math.Max(0, glyphId), _entries.Length - 1)];

        internal void Validate(OfficeOpenTypeItemVariationStore store) {
            for (int index = 0; index < _entries.Length; index++) {
                store.ValidateIndex(_entries[index].Outer, _entries[index].Inner);
            }
        }
    }

    private readonly struct DeltaSetIndex {
        internal DeltaSetIndex(int outer, int inner) {
            Outer = outer;
            Inner = inner;
        }

        internal int Outer { get; }
        internal int Inner { get; }
    }
}
