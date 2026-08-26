using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>Validated fvar/avar axis model shared by TrueType and CFF2 outline programs.</summary>
internal sealed class OfficeFontVariationModel {
    private readonly Axis[] _axes;
    private readonly double[] _normalizedCoordinates;
    private readonly IReadOnlyDictionary<string, float> _designCoordinates;

    private OfficeFontVariationModel(
        Axis[] axes,
        double[] normalizedCoordinates,
        IReadOnlyDictionary<string, float> designCoordinates,
        string identity) {
        _axes = axes;
        _normalizedCoordinates = normalizedCoordinates;
        _designCoordinates = designCoordinates;
        Identity = identity;
    }

    internal static OfficeFontVariationModel None { get; } = new(
        Array.Empty<Axis>(),
        Array.Empty<double>(),
        new ReadOnlyDictionary<string, float>(new Dictionary<string, float>(StringComparer.Ordinal)),
        "static");

    internal bool IsVariable => _axes.Length > 0;
    internal int AxisCount => _axes.Length;
    internal string Identity { get; }
    internal IReadOnlyList<double> NormalizedCoordinates => _normalizedCoordinates;
    internal IReadOnlyDictionary<string, float> DesignCoordinates => _designCoordinates;

    internal static OfficeFontVariationModel Create(
        OfficeOpenTypeReader reader,
        IReadOnlyDictionary<string, float>? requestedValues) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        if (!reader.TryGetTable("fvar", out int offset, out int length)) {
            if (requestedValues != null && requestedValues.Count > 0) {
                throw new ArgumentException("Variable-font axes were supplied for a static font.", nameof(requestedValues));
            }
            return None;
        }
        if (length < 16) throw new InvalidDataException("The OpenType fvar table is truncated.");
        if (reader.ReadUInt16(offset) != 1 || reader.ReadUInt16(offset + 6) != 2) {
            throw new InvalidDataException("The OpenType fvar header is invalid.");
        }
        int axesOffset = checked(offset + reader.ReadUInt16(offset + 4));
        int axisCount = reader.ReadUInt16(offset + 8);
        int axisSize = reader.ReadUInt16(offset + 10);
        int instanceCount = reader.ReadUInt16(offset + 12);
        int instanceSize = reader.ReadUInt16(offset + 14);
        if (axisCount == 0) {
            if (requestedValues != null && requestedValues.Count > 0) {
                throw new ArgumentException("Variable-font axes were supplied for a static font.", nameof(requestedValues));
            }
            return None;
        }
        int minimumInstanceSize = checked(axisCount * 4 + 4);
        if (axisCount > 64 || axisSize < 20 || instanceCount > 4096 || instanceSize < minimumInstanceSize
            || axesOffset < offset + 16
            || axesOffset > offset + length - checked(axisCount * axisSize)) {
            throw new InvalidDataException("The OpenType fvar axis directory is invalid.");
        }
        int instancesOffset = checked(axesOffset + axisCount * axisSize);
        if (instancesOffset > offset + length - checked(instanceCount * instanceSize)) {
            throw new InvalidDataException("The OpenType fvar instance directory is invalid.");
        }

        var axes = new Axis[axisCount];
        var axisIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int index = 0; index < axisCount; index++) {
            int record = axesOffset + index * axisSize;
            string tag = ReadTag(reader, record);
            double minimum = reader.ReadFixed16_16(record + 4);
            double defaultValue = reader.ReadFixed16_16(record + 8);
            double maximum = reader.ReadFixed16_16(record + 12);
            if (!IsFinite(minimum) || !IsFinite(defaultValue) || !IsFinite(maximum)
                || minimum > defaultValue || defaultValue > maximum || axisIndexes.ContainsKey(tag)) {
                throw new InvalidDataException("The OpenType fvar axis data is invalid.");
            }
            axes[index] = new Axis(tag, minimum, defaultValue, maximum);
            axisIndexes.Add(tag, index);
        }

        var values = new double[axisCount];
        for (int index = 0; index < axisCount; index++) values[index] = axes[index].Default;
        if (requestedValues != null) {
            foreach (KeyValuePair<string, float> requested in requestedValues) {
                ValidateTag(requested.Key);
                if (float.IsNaN(requested.Value) || float.IsInfinity(requested.Value)) {
                    throw new ArgumentOutOfRangeException(nameof(requestedValues), "Variable-font axis values must be finite.");
                }
                if (!axisIndexes.TryGetValue(requested.Key, out int axisIndex)) {
                    throw new ArgumentException("Variable-font axis '" + requested.Key + "' is not defined by the font.", nameof(requestedValues));
                }
                Axis axis = axes[axisIndex];
                values[axisIndex] = Math.Max(axis.Minimum, Math.Min(axis.Maximum, requested.Value));
            }
        }

        var normalized = new double[axisCount];
        for (int index = 0; index < axisCount; index++) normalized[index] = Normalize(axes[index], values[index]);
        ApplyAvar(reader, axes, normalized);
        var identity = new StringBuilder();
        var designCoordinates = new Dictionary<string, float>(axisCount, StringComparer.Ordinal);
        for (int index = 0; index < axisCount; index++) {
            if (index > 0) identity.Append(';');
            identity.Append(axes[index].Tag)
                .Append('=')
                .Append(values[index].ToString("R", CultureInfo.InvariantCulture));
            designCoordinates.Add(axes[index].Tag, checked((float)values[index]));
        }
        return new OfficeFontVariationModel(
            axes,
            normalized,
            new ReadOnlyDictionary<string, float>(designCoordinates),
            identity.ToString());
    }

    private static void ApplyAvar(
        OfficeOpenTypeReader reader,
        Axis[] axes,
        double[] normalized) {
        if (!reader.TryGetTable("avar", out int offset, out int length)) return;
        if (length < 8 || reader.ReadUInt16(offset) != 1 || reader.ReadUInt16(offset + 2) != 0
            || reader.ReadUInt16(offset + 4) != 0 || reader.ReadUInt16(offset + 6) != axes.Length) {
            throw new InvalidDataException("The OpenType avar header is invalid.");
        }
        int cursor = offset + 8;
        int end = checked(offset + length);
        for (int axisIndex = 0; axisIndex < axes.Length; axisIndex++) {
            if (cursor > end - 2) throw new InvalidDataException("The OpenType avar segment map is truncated.");
            int mapCount = reader.ReadUInt16(cursor);
            cursor += 2;
            if (mapCount > 4096 || cursor > end - mapCount * 4) {
                throw new InvalidDataException("The OpenType avar segment map is invalid.");
            }
            var from = new double[mapCount];
            var to = new double[mapCount];
            bool validMap = mapCount >= 3;
            for (int index = 0; index < mapCount; index++) {
                from[index] = reader.ReadF2Dot14(cursor);
                to[index] = reader.ReadF2Dot14(cursor + 2);
                cursor += 4;
                if (from[index] < -1D || from[index] > 1D || to[index] < -1D || to[index] > 1D
                    || index > 0 && (from[index] <= from[index - 1] || to[index] < to[index - 1])) validMap = false;
            }
            validMap &= ContainsMapping(from, to, -1D, -1D)
                && ContainsMapping(from, to, 0D, 0D)
                && ContainsMapping(from, to, 1D, 1D);
            if (!validMap) throw new InvalidDataException("The OpenType avar segment map is invalid.");
            normalized[axisIndex] = MapSegment(normalized[axisIndex], from, to);
        }
        if (cursor != end) throw new InvalidDataException("The OpenType avar table contains trailing data.");
    }

    private static bool ContainsMapping(double[] from, double[] to, double source, double target) {
        for (int index = 0; index < from.Length; index++) {
            if (from[index] == source && to[index] == target) return true;
        }
        return false;
    }

    private static double MapSegment(double value, double[] from, double[] to) {
        if (value <= from[0]) return to[0];
        if (value >= from[from.Length - 1]) return to[to.Length - 1];
        for (int index = 1; index < from.Length; index++) {
            if (value > from[index]) continue;
            double span = from[index] - from[index - 1];
            if (span <= 0D) throw new InvalidDataException("The OpenType avar segment span is invalid.");
            double ratio = (value - from[index - 1]) / span;
            return to[index - 1] + ((to[index] - to[index - 1]) * ratio);
        }
        return value;
    }

    private static double Normalize(Axis axis, double value) {
        if (value == axis.Default) return 0D;
        if (value < axis.Default) {
            double range = axis.Default - axis.Minimum;
            return range <= 0D ? 0D : Math.Max(-1D, (value - axis.Default) / range);
        }
        double upperRange = axis.Maximum - axis.Default;
        return upperRange <= 0D ? 0D : Math.Min(1D, (value - axis.Default) / upperRange);
    }

    private static string ReadTag(OfficeOpenTypeReader reader, int offset) {
        uint value = reader.ReadUInt32(offset);
        var characters = new[] {
            (char)((value >> 24) & 0xFF),
            (char)((value >> 16) & 0xFF),
            (char)((value >> 8) & 0xFF),
            (char)(value & 0xFF)
        };
        string tag = new string(characters);
        ValidateTag(tag);
        return tag;
    }

    private static void ValidateTag(string? tag) {
        if (tag == null || tag.Length != 4) throw new ArgumentException("Variable-font axis tags must contain exactly four characters.");
        for (int index = 0; index < tag.Length; index++) {
            if (tag[index] < 0x20 || tag[index] > 0x7E) throw new ArgumentException("Variable-font axis tags must contain printable ASCII characters.");
        }
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private readonly struct Axis {
        internal Axis(string tag, double minimum, double defaultValue, double maximum) {
            Tag = tag;
            Minimum = minimum;
            Default = defaultValue;
            Maximum = maximum;
        }

        internal string Tag { get; }
        internal double Minimum { get; }
        internal double Default { get; }
        internal double Maximum { get; }
    }
}
