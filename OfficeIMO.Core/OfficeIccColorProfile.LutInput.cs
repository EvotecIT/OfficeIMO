using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeIccColorProfile {
    private static bool TryReadLutTransform(
        byte[] bytes,
        TagRange range,
        int expectedInputChannels,
        bool pcsIsLab,
        out LutTransform transform) {
        transform = null!;
        if (range.Length < 52) return false;
        uint type = ReadUInt32(bytes, range.Offset);
        int precision = type == Lut8TypeSignature ? 1 : type == Lut16TypeSignature ? 2 : 0;
        if (precision == 0 || (precision == 1 && !pcsIsLab) ||
            bytes[range.Offset + 4] != 0 || bytes[range.Offset + 5] != 0 ||
            bytes[range.Offset + 6] != 0 || bytes[range.Offset + 7] != 0 || bytes[range.Offset + 11] != 0) return false;
        int inputChannels = bytes[range.Offset + 8];
        int outputChannels = bytes[range.Offset + 9];
        int gridPoints = bytes[range.Offset + 10];
        if (inputChannels != expectedInputChannels || outputChannels != 3 ||
            inputChannels is < 3 or > 8 || gridPoints is < 2 or > 33 ||
            !HasIdentityLutMatrix(bytes, range.Offset + 12)) return false;

        int inputEntries = precision == 1 ? 256 : ReadUInt16(bytes, range.Offset + 48);
        int outputEntries = precision == 1 ? 256 : ReadUInt16(bytes, range.Offset + 50);
        int tableOffset = precision == 1 ? 48 : 52;
        if (inputEntries is < 2 or > 4096 || outputEntries is < 2 or > 4096) return false;
        long gridSampleCount = 1;
        for (int channel = 0; channel < inputChannels; channel++) {
            gridSampleCount *= gridPoints;
            if (gridSampleCount > int.MaxValue) return false;
        }
        long inputBytes = (long)inputChannels * inputEntries * precision;
        long clutBytes = gridSampleCount * outputChannels * precision;
        long outputBytes = (long)outputChannels * outputEntries * precision;
        long requiredLength = tableOffset + inputBytes + clutBytes + outputBytes;
        if (requiredLength != range.Length || requiredLength > int.MaxValue) return false;

        var payload = new byte[(int)requiredLength];
        Buffer.BlockCopy(bytes, range.Offset, payload, 0, payload.Length);
        transform = new LutTransform(
            payload,
            inputChannels,
            outputChannels,
            gridPoints,
            inputEntries,
            outputEntries,
            precision,
            tableOffset,
            checked(tableOffset + (int)inputBytes),
            checked(tableOffset + (int)inputBytes + (int)clutBytes),
            pcsIsLab);
        return true;
    }

    private static bool HasIdentityLutMatrix(byte[] bytes, int offset) {
        for (int row = 0; row < 3; row++) {
            for (int column = 0; column < 3; column++) {
                int expected = row == column ? 65536 : 0;
                if (unchecked((int)ReadUInt32(bytes, offset + (row * 3 + column) * 4)) != expected) return false;
            }
        }
        return true;
    }

    private sealed class LutTransform : IDeviceToPcsTransform {
        private const double PcsXyzScale = 65535D / 32768D;
        private readonly byte[] _payload;
        private readonly int _inputChannels;
        private readonly int _outputChannels;
        private readonly int _gridPoints;
        private readonly int _inputEntries;
        private readonly int _outputEntries;
        private readonly int _precision;
        private readonly int _inputOffset;
        private readonly int _clutOffset;
        private readonly int _outputOffset;
        private readonly bool _pcsIsLab;

        internal LutTransform(
            byte[] payload,
            int inputChannels,
            int outputChannels,
            int gridPoints,
            int inputEntries,
            int outputEntries,
            int precision,
            int inputOffset,
            int clutOffset,
            int outputOffset,
            bool pcsIsLab) {
            _payload = payload;
            _inputChannels = inputChannels;
            _outputChannels = outputChannels;
            _gridPoints = gridPoints;
            _inputEntries = inputEntries;
            _outputEntries = outputEntries;
            _precision = precision;
            _inputOffset = inputOffset;
            _clutOffset = clutOffset;
            _outputOffset = outputOffset;
            _pcsIsLab = pcsIsLab;
        }

        public long RetainedByteCount => checked(96L + _payload.LongLength);

        public bool TryTransform(IReadOnlyList<double> components, XyzValue whitePoint, out XyzValue pcsXyz) {
            pcsXyz = default;
            if (components.Count < _inputChannels) return false;
            return Transform(new DeviceComponentValues(components, _inputChannels), whitePoint, out pcsXyz);
        }

        public bool TryTransform(DeviceComponentValues components, XyzValue whitePoint, out XyzValue pcsXyz) {
            pcsXyz = default;
            return components.Count >= _inputChannels && Transform(components, whitePoint, out pcsXyz);
        }

        private bool Transform(DeviceComponentValues components, XyzValue whitePoint, out XyzValue pcsXyz) {
            var input = new DeviceComponentValues(_inputChannels,
                LookupInput(components[0], 0),
                LookupInput(components[1], 1),
                LookupInput(components[2], 2),
                _inputChannels > 3 ? LookupInput(components[3], 3) : 0D,
                _inputChannels > 4 ? LookupInput(components[4], 4) : 0D,
                _inputChannels > 5 ? LookupInput(components[5], 5) : 0D,
                _inputChannels > 6 ? LookupInput(components[6], 6) : 0D,
                _inputChannels > 7 ? LookupInput(components[7], 7) : 0D);
            InterpolateClut(input, out double output0, out double output1, out double output2);
            output0 = LookupTable(_outputOffset, _outputEntries, output0);
            output1 = LookupTable(_outputOffset + _outputEntries * _precision, _outputEntries, output1);
            output2 = LookupTable(_outputOffset + 2 * _outputEntries * _precision, _outputEntries, output2);
            if (_pcsIsLab) {
                double lightness = _precision == 1 ? output0 * 100D : output0 * (65535D / 65280D) * 100D;
                double a = _precision == 1 ? output1 * 255D - 128D : output1 * (65535D / 256D) - 128D;
                double b = _precision == 1 ? output2 * 255D - 128D : output2 * (65535D / 256D) - 128D;
                OfficeColorSpaceConverter.ConvertLabToXyz(
                    lightness,
                    a,
                    b,
                    whitePoint.X,
                    whitePoint.Y,
                    whitePoint.Z,
                    out double x,
                    out double y,
                    out double z);
                pcsXyz = new XyzValue(x, y, z);
            } else {
                pcsXyz = new XyzValue(
                    output0 * PcsXyzScale,
                    output1 * PcsXyzScale,
                    output2 * PcsXyzScale);
            }
            return true;
        }

        private double LookupInput(double component, int channel) =>
            LookupTable(
                _inputOffset + channel * _inputEntries * _precision,
                _inputEntries,
                Clamp01(component));

        private double LookupTable(int offset, int entries, double value) {
            double position = Clamp01(value) * (entries - 1);
            int lower = (int)Math.Floor(position);
            if (lower >= entries - 1) return ReadNormalized(offset + (entries - 1) * _precision);
            double fraction = position - lower;
            double left = ReadNormalized(offset + lower * _precision);
            double right = ReadNormalized(offset + (lower + 1) * _precision);
            return left + (right - left) * fraction;
        }

        private void InterpolateClut(DeviceComponentValues input,
            out double output0, out double output1, out double output2) {
            var positions = new DeviceComponentValues(_inputChannels,
                Clamp01(input[0]) * (_gridPoints - 1),
                Clamp01(input[1]) * (_gridPoints - 1),
                Clamp01(input[2]) * (_gridPoints - 1),
                _inputChannels > 3 ? Clamp01(input[3]) * (_gridPoints - 1) : 0D,
                _inputChannels > 4 ? Clamp01(input[4]) * (_gridPoints - 1) : 0D,
                _inputChannels > 5 ? Clamp01(input[5]) * (_gridPoints - 1) : 0D,
                _inputChannels > 6 ? Clamp01(input[6]) * (_gridPoints - 1) : 0D,
                _inputChannels > 7 ? Clamp01(input[7]) * (_gridPoints - 1) : 0D);
            output0 = 0D;
            output1 = 0D;
            output2 = 0D;
            int cornerCount = 1 << _inputChannels;
            for (int corner = 0; corner < cornerCount; corner++) {
                double weight = 1D;
                int gridIndex = 0;
                for (int channel = 0; channel < _inputChannels; channel++) {
                    bool upper = (corner & (1 << channel)) != 0;
                    int lower = Math.Min((int)positions[channel], _gridPoints - 2);
                    double fraction = positions[channel] - lower;
                    weight *= upper ? fraction : 1D - fraction;
                    gridIndex = gridIndex * _gridPoints + lower + (upper ? 1 : 0);
                }
                if (weight == 0D) continue;
                int valueOffset = _clutOffset + gridIndex * _outputChannels * _precision;
                output0 += ReadNormalized(valueOffset) * weight;
                output1 += ReadNormalized(valueOffset + _precision) * weight;
                output2 += ReadNormalized(valueOffset + 2 * _precision) * weight;
            }
        }

        private double ReadNormalized(int offset) => _precision == 1
            ? _payload[offset] / 255D
            : ReadUInt16(_payload, offset) / 65535D;
    }
}
