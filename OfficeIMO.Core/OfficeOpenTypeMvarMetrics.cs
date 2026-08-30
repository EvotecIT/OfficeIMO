using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Validated MVAR evaluator for variable-font line metrics.</summary>
internal sealed class OfficeOpenTypeMvarMetrics {
    private const uint HorizontalAscenderTag = 0x68617363; // hasc
    private const uint HorizontalDescenderTag = 0x68647363; // hdsc
    private const uint HorizontalLineGapTag = 0x686C6770; // hlgp
    private readonly OfficeOpenTypeItemVariationStore _store;
    private readonly Dictionary<uint, ValueRecord> _records;

    private OfficeOpenTypeMvarMetrics(
        OfficeOpenTypeItemVariationStore store,
        Dictionary<uint, ValueRecord> records) {
        _store = store;
        _records = records;
    }

    internal static OfficeOpenTypeMvarMetrics? TryParse(
        OfficeOpenTypeReader reader,
        OfficeFontVariationModel model) {
        if (!reader.TryGetTable("MVAR", out int offset, out int length)) return null;
        return Parse(reader, offset, length, model);
    }

    internal static OfficeOpenTypeMvarMetrics Parse(
        OfficeOpenTypeReader reader,
        int offset,
        int length,
        OfficeFontVariationModel model) {
        if (length < 12 || offset < 0 || offset > reader.Data.Length - length) {
            throw new InvalidDataException("The MVAR table is truncated.");
        }
        int end = checked(offset + length);
        if (reader.ReadUInt16(offset) != 1 || reader.ReadUInt16(offset + 2) != 0
            || reader.ReadUInt16(offset + 4) != 0) {
            throw new InvalidDataException("The MVAR table header is invalid.");
        }
        int recordSize = reader.ReadUInt16(offset + 6);
        int recordCount = reader.ReadUInt16(offset + 8);
        int storeRelative = reader.ReadUInt16(offset + 10);
        if (recordSize < 8 || recordCount > 4096) {
            throw new InvalidDataException("The MVAR value-record directory is invalid.");
        }
        int recordsOffset = offset + 12;
        int recordsLength = checked(recordSize * recordCount);
        int recordsEnd = checked(recordsOffset + recordsLength);
        int storeOffset = checked(offset + storeRelative);
        if (recordsEnd > end || storeRelative == 0 || storeOffset < recordsEnd || storeOffset >= end) {
            throw new InvalidDataException("The MVAR ItemVariationStore offset is invalid.");
        }

        var records = new Dictionary<uint, ValueRecord>();
        var allRecordIndexes = new List<ValueRecord>(recordCount);
        uint previousTag = 0;
        for (int index = 0; index < recordCount; index++) {
            int recordOffset = checked(recordsOffset + index * recordSize);
            uint tag = reader.ReadUInt32(recordOffset);
            if (index > 0 && tag <= previousTag) {
                throw new InvalidDataException("The MVAR value records are not strictly tag-sorted.");
            }
            previousTag = tag;
            var valueRecord = new ValueRecord(
                reader.ReadUInt16(recordOffset + 4),
                reader.ReadUInt16(recordOffset + 6));
            allRecordIndexes.Add(valueRecord);
            if (tag == HorizontalAscenderTag || tag == HorizontalDescenderTag || tag == HorizontalLineGapTag) {
                records.Add(tag, valueRecord);
            }
        }
        OfficeOpenTypeItemVariationStore store = OfficeOpenTypeItemVariationStore.Parse(
            reader,
            storeOffset,
            end,
            model);
        for (int index = 0; index < allRecordIndexes.Count; index++) {
            ValueRecord record = allRecordIndexes[index];
            store.ValidateIndex(record.OuterIndex, record.InnerIndex);
        }
        return new OfficeOpenTypeMvarMetrics(store, records);
    }

    internal int HorizontalAscenderDelta => Evaluate(HorizontalAscenderTag);
    internal int HorizontalDescenderDelta => Evaluate(HorizontalDescenderTag);
    internal int HorizontalLineGapDelta => Evaluate(HorizontalLineGapTag);

    private int Evaluate(uint tag) => _records.TryGetValue(tag, out ValueRecord record)
        ? _store.Evaluate(record.OuterIndex, record.InnerIndex)
        : 0;

    private readonly struct ValueRecord {
        internal ValueRecord(int outerIndex, int innerIndex) {
            OuterIndex = outerIndex;
            InnerIndex = innerIndex;
        }

        internal int OuterIndex { get; }
        internal int InnerIndex { get; }
    }
}
