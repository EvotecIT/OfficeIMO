using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Validated OpenType ItemVariationStore evaluator used by variable metrics.</summary>
internal sealed class OfficeOpenTypeItemVariationStore {
    private readonly OfficeOpenTypeReader _reader;
    private readonly DataSet[] _dataSets;
    private readonly double[] _regionScalars;

    private OfficeOpenTypeItemVariationStore(
        OfficeOpenTypeReader reader,
        DataSet[] dataSets,
        double[] regionScalars) {
        _reader = reader;
        _dataSets = dataSets;
        _regionScalars = regionScalars;
    }

    internal static OfficeOpenTypeItemVariationStore Parse(
        OfficeOpenTypeReader reader,
        int offset,
        int end,
        OfficeFontVariationModel variations) {
        if (offset < 0 || offset > end - 8 || reader.ReadUInt16(offset) != 1) {
            throw new InvalidDataException("The OpenType ItemVariationStore header is invalid.");
        }
        uint regionRelativeValue = reader.ReadUInt32(offset + 2);
        int dataCount = reader.ReadUInt16(offset + 6);
        if (regionRelativeValue > int.MaxValue || dataCount > 4096 || offset + 8 > end - checked(dataCount * 4)) {
            throw new InvalidDataException("The OpenType ItemVariationStore directory is invalid.");
        }
        int regionList = checked(offset + (int)regionRelativeValue);
        if (regionList < offset || regionList > end - 4) {
            throw new InvalidDataException("The OpenType VariationRegionList offset is invalid.");
        }
        int axisCount = reader.ReadUInt16(regionList);
        int regionCount = reader.ReadUInt16(regionList + 2);
        if (axisCount != variations.AxisCount || regionCount > 32768
            || regionList + 4 > end - checked(axisCount * regionCount * 6)) {
            throw new InvalidDataException("The OpenType VariationRegionList dimensions are invalid.");
        }
        var regionScalars = new double[regionCount];
        int cursor = regionList + 4;
        for (int region = 0; region < regionCount; region++) {
            double scalar = 1D;
            for (int axis = 0; axis < axisCount; axis++) {
                double start = reader.ReadF2Dot14(cursor);
                double peak = reader.ReadF2Dot14(cursor + 2);
                double finish = reader.ReadF2Dot14(cursor + 4);
                cursor += 6;
                scalar *= OfficeOpenTypeVariationRegion.CalculateScalar(
                    variations.NormalizedCoordinates[axis],
                    start,
                    peak,
                    finish);
            }
            regionScalars[region] = scalar;
        }

        var dataSets = new DataSet[dataCount];
        var dataSetsByOffset = new Dictionary<int, DataSet>();
        int persistentRegionIndexCount = 0;
        int maximumPersistentRegionIndexes = Math.Min(
            1_000_000,
            Math.Max(64, reader.Data.Length <= int.MaxValue / 2 ? reader.Data.Length * 2 : 1_000_000));
        for (int dataIndex = 0; dataIndex < dataCount; dataIndex++) {
            uint relativeValue = reader.ReadUInt32(offset + 8 + dataIndex * 4);
            if (relativeValue == 0) {
                dataSets[dataIndex] = default;
                continue;
            }
            if (relativeValue > int.MaxValue) throw new InvalidDataException("An ItemVariationData offset is invalid.");
            int dataOffset = checked(offset + (int)relativeValue);
            if (dataOffset < offset || dataOffset > end - 6) throw new InvalidDataException("An ItemVariationData header is truncated.");
            if (dataSetsByOffset.TryGetValue(dataOffset, out DataSet existingDataSet)) {
                dataSets[dataIndex] = existingDataSet;
                continue;
            }
            int itemCount = reader.ReadUInt16(dataOffset);
            int wordDeltaCountValue = reader.ReadUInt16(dataOffset + 2);
            bool longWords = (wordDeltaCountValue & 0x8000) != 0;
            int wordDeltaCount = wordDeltaCountValue & 0x7FFF;
            int regionIndexCount = reader.ReadUInt16(dataOffset + 4);
            if (itemCount > 65535 || wordDeltaCount > regionIndexCount || regionIndexCount > regionCount
                || dataOffset + 6 > end - checked(regionIndexCount * 2)) {
                throw new InvalidDataException("An ItemVariationData directory is invalid.");
            }
            if (regionIndexCount > maximumPersistentRegionIndexes - persistentRegionIndexCount) {
                throw new InvalidDataException("OpenType variation metadata exceeds the bounded allocation budget.");
            }
            persistentRegionIndexCount += regionIndexCount;
            var regionIndexes = new ushort[regionIndexCount];
            cursor = dataOffset + 6;
            for (int index = 0; index < regionIndexCount; index++) {
                int regionIndex = reader.ReadUInt16(cursor);
                cursor += 2;
                if (regionIndex >= regionCount) throw new InvalidDataException("An ItemVariationData region index is invalid.");
                regionIndexes[index] = (ushort)regionIndex;
            }
            int largeSize = longWords ? 4 : 2;
            int smallSize = longWords ? 2 : 1;
            int rowSize = checked(wordDeltaCount * largeSize + (regionIndexCount - wordDeltaCount) * smallSize);
            if (cursor > end - checked(itemCount * rowSize)) throw new InvalidDataException("An ItemVariationData delta array is truncated.");
            var dataSet = new DataSet(cursor, itemCount, wordDeltaCount, regionIndexes, longWords, rowSize);
            dataSetsByOffset.Add(dataOffset, dataSet);
            dataSets[dataIndex] = dataSet;
        }
        return new OfficeOpenTypeItemVariationStore(reader, dataSets, regionScalars);
    }

    internal int Evaluate(int outerIndex, int innerIndex) {
        if (outerIndex < 0 || outerIndex >= _dataSets.Length) return 0;
        DataSet dataSet = _dataSets[outerIndex];
        if (!dataSet.IsPresent) return 0;
        if (innerIndex < 0 || innerIndex >= dataSet.ItemCount) return 0;
        int cursor = checked(dataSet.Offset + innerIndex * dataSet.RowSize);
        double value = 0D;
        for (int index = 0; index < dataSet.RegionIndexes.Length; index++) {
            int delta;
            if (index < dataSet.WordDeltaCount) {
                if (dataSet.LongWords) {
                    delta = _reader.ReadInt32(cursor);
                    cursor += 4;
                } else {
                    delta = _reader.ReadInt16(cursor);
                    cursor += 2;
                }
            } else if (dataSet.LongWords) {
                delta = _reader.ReadInt16(cursor);
                cursor += 2;
            } else {
                _reader.EnsureAvailable(cursor, 1);
                delta = unchecked((sbyte)_reader.Data[cursor++]);
            }
            value += delta * _regionScalars[dataSet.RegionIndexes[index]];
        }
        return checked((int)Math.Round(value, MidpointRounding.ToEven));
    }

    private readonly struct DataSet {
        internal DataSet(int offset, int itemCount, int wordDeltaCount, ushort[] regionIndexes, bool longWords, int rowSize) {
            IsPresent = true;
            Offset = offset;
            ItemCount = itemCount;
            WordDeltaCount = wordDeltaCount;
            RegionIndexes = regionIndexes;
            LongWords = longWords;
            RowSize = rowSize;
        }

        internal bool IsPresent { get; }
        internal int Offset { get; }
        internal int ItemCount { get; }
        internal int WordDeltaCount { get; }
        internal ushort[] RegionIndexes { get; }
        internal bool LongWords { get; }
        internal int RowSize { get; }
    }
}
