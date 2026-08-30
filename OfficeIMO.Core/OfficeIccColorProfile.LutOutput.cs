using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeIccColorProfile {
    private const int MaximumLegacyOutputLutGridPoints = 65;
    private const int MaximumLegacyOutputLutPayloadBytes = 64 * 1024 * 1024;

    private static bool TryReadLutPcsToDeviceTransform(
        byte[] bytes,
        TagRange range,
        int expectedOutputChannels,
        bool pcsIsLab,
        out LutPcsToDeviceTransform transform) {
        transform = null!;
        if (range.Length < 52) return false;

        uint type = ReadUInt32(bytes, range.Offset);
        int precision = type == Lut8TypeSignature ? 1 : type == Lut16TypeSignature ? 2 : 0;
        if (precision == 0 ||
            !AreZero(bytes, range.Offset + 4, 4) || bytes[range.Offset + 11] != 0) {
            return false;
        }

        int inputChannels = bytes[range.Offset + 8];
        int outputChannels = bytes[range.Offset + 9];
        int gridPoints = bytes[range.Offset + 10];
        if (inputChannels != 3 || outputChannels != expectedOutputChannels ||
            outputChannels is < 3 or > 4 || gridPoints is < 2 or > MaximumLegacyOutputLutGridPoints ||
            pcsIsLab && !HasIdentityLutMatrix(bytes, range.Offset + 12)) {
            return false;
        }
        double[] matrix = ReadLegacyLutMatrix(bytes, range.Offset + 12);

        int inputEntries = precision == 1 ? 256 : ReadUInt16(bytes, range.Offset + 48);
        int outputEntries = precision == 1 ? 256 : ReadUInt16(bytes, range.Offset + 50);
        int tableOffset = precision == 1 ? 48 : 52;
        if (inputEntries is < 2 or > 4096 || outputEntries is < 2 or > 4096) return false;

        long gridSampleCount = (long)gridPoints * gridPoints * gridPoints;
        long inputBytes = (long)inputChannels * inputEntries * precision;
        long clutBytes = gridSampleCount * outputChannels * precision;
        long outputBytes = (long)outputChannels * outputEntries * precision;
        long requiredLength = tableOffset + inputBytes + clutBytes + outputBytes;
        if (requiredLength != range.Length || requiredLength > MaximumLegacyOutputLutPayloadBytes) return false;

        var payload = new byte[(int)requiredLength];
        Buffer.BlockCopy(bytes, range.Offset, payload, 0, payload.Length);
        transform = new LutPcsToDeviceTransform(
            payload,
            outputChannels,
            gridPoints,
            inputEntries,
            outputEntries,
            precision,
            tableOffset,
            checked(tableOffset + (int)inputBytes),
            checked(tableOffset + (int)inputBytes + (int)clutBytes),
            pcsIsLab,
            matrix);
        return true;
    }

    private static double[] ReadLegacyLutMatrix(byte[] bytes, int offset) {
        var matrix = new double[9];
        for (int index = 0; index < matrix.Length; index++) {
            matrix[index] = ReadS15Fixed16(bytes, offset + index * 4);
        }
        return matrix;
    }

    private sealed class LutPcsToDeviceTransform : IPcsToDeviceTransform {
        private const double PcsXyzScale = 65535D / 32768D;
        private readonly byte[] _payload;
        private readonly int _outputChannels;
        private readonly int _gridPoints;
        private readonly int _inputEntries;
        private readonly int _outputEntries;
        private readonly int _precision;
        private readonly int _inputOffset;
        private readonly int _clutOffset;
        private readonly int _outputOffset;
        private readonly bool _pcsIsLab;
        private readonly double[] _matrix;

        internal LutPcsToDeviceTransform(
            byte[] payload,
            int outputChannels,
            int gridPoints,
            int inputEntries,
            int outputEntries,
            int precision,
            int inputOffset,
            int clutOffset,
            int outputOffset,
            bool pcsIsLab,
            double[] matrix) {
            _payload = payload;
            _outputChannels = outputChannels;
            _gridPoints = gridPoints;
            _inputEntries = inputEntries;
            _outputEntries = outputEntries;
            _precision = precision;
            _inputOffset = inputOffset;
            _clutOffset = clutOffset;
            _outputOffset = outputOffset;
            _pcsIsLab = pcsIsLab;
            _matrix = matrix;
        }

        public long RetainedByteCount => checked(96L + _payload.LongLength + (_matrix.LongLength * sizeof(double)));

        public bool TryTransform(XyzValue pcsXyz, XyzValue whitePoint, out DeviceComponentValues components) {
            double input0;
            double input1;
            double input2;
            if (_pcsIsLab) {
                ConvertXyzToLab(pcsXyz, whitePoint, out double lightness, out double a, out double b);
                if (_precision == 1) {
                    input0 = lightness / 100D;
                    input1 = (a + 128D) / 255D;
                    input2 = (b + 128D) / 255D;
                } else {
                    input0 = (lightness / 100D) * (65280D / 65535D);
                    input1 = (a + 128D) * (256D / 65535D);
                    input2 = (b + 128D) * (256D / 65535D);
                }
            } else {
                input0 = pcsXyz.X / PcsXyzScale;
                input1 = pcsXyz.Y / PcsXyzScale;
                input2 = pcsXyz.Z / PcsXyzScale;
                double transformed0 = (_matrix[0] * input0) + (_matrix[1] * input1) + (_matrix[2] * input2);
                double transformed1 = (_matrix[3] * input0) + (_matrix[4] * input1) + (_matrix[5] * input2);
                double transformed2 = (_matrix[6] * input0) + (_matrix[7] * input1) + (_matrix[8] * input2);
                input0 = transformed0;
                input1 = transformed1;
                input2 = transformed2;
            }

            input0 = LookupTable(_inputOffset, _inputEntries, input0);
            input1 = LookupTable(_inputOffset + _inputEntries * _precision, _inputEntries, input1);
            input2 = LookupTable(_inputOffset + 2 * _inputEntries * _precision, _inputEntries, input2);
            InterpolateClut(input0, input1, input2, out double output0, out double output1, out double output2, out double output3);
            output0 = LookupOutput(output0, 0);
            output1 = LookupOutput(output1, 1);
            output2 = LookupOutput(output2, 2);
            output3 = _outputChannels == 4 ? LookupOutput(output3, 3) : 0D;
            components = new DeviceComponentValues(_outputChannels, output0, output1, output2, output3);
            return true;
        }

        private double LookupOutput(double value, int channel) =>
            LookupTable(
                _outputOffset + channel * _outputEntries * _precision,
                _outputEntries,
                value);

        private double LookupTable(int offset, int entries, double value) {
            double position = Clamp01(value) * (entries - 1);
            int lower = (int)Math.Floor(position);
            if (lower >= entries - 1) return ReadNormalized(offset + (entries - 1) * _precision);
            double fraction = position - lower;
            double lowerValue = ReadNormalized(offset + lower * _precision);
            double upperValue = ReadNormalized(offset + (lower + 1) * _precision);
            return lowerValue + ((upperValue - lowerValue) * fraction);
        }

        private void InterpolateClut(
            double input0,
            double input1,
            double input2,
            out double output0,
            out double output1,
            out double output2,
            out double output3) {
            GetGridPosition(input0, out int lower0, out double fraction0);
            GetGridPosition(input1, out int lower1, out double fraction1);
            GetGridPosition(input2, out int lower2, out double fraction2);
            output0 = 0D;
            output1 = 0D;
            output2 = 0D;
            output3 = 0D;
            for (int corner = 0; corner < 8; corner++) {
                bool upper0 = (corner & 4) != 0;
                bool upper1 = (corner & 2) != 0;
                bool upper2 = (corner & 1) != 0;
                double weight =
                    (upper0 ? fraction0 : 1D - fraction0) *
                    (upper1 ? fraction1 : 1D - fraction1) *
                    (upper2 ? fraction2 : 1D - fraction2);
                if (weight == 0D) continue;
                int gridIndex =
                    (((lower0 + (upper0 ? 1 : 0)) * _gridPoints) + lower1 + (upper1 ? 1 : 0)) * _gridPoints +
                    lower2 + (upper2 ? 1 : 0);
                int offset = _clutOffset + gridIndex * _outputChannels * _precision;
                output0 += ReadNormalized(offset) * weight;
                output1 += ReadNormalized(offset + _precision) * weight;
                output2 += ReadNormalized(offset + 2 * _precision) * weight;
                if (_outputChannels == 4) output3 += ReadNormalized(offset + 3 * _precision) * weight;
            }
        }

        private void GetGridPosition(double input, out int lower, out double fraction) {
            double position = Clamp01(input) * (_gridPoints - 1);
            lower = Math.Min((int)Math.Floor(position), _gridPoints - 2);
            fraction = position - lower;
        }

        private double ReadNormalized(int offset) => _precision == 1
            ? _payload[offset] / 255D
            : ReadUInt16(_payload, offset) / 65535D;

        private static void ConvertXyzToLab(
            XyzValue xyz,
            XyzValue whitePoint,
            out double lightness,
            out double a,
            out double b) {
            double fx = LabFunction(xyz.X / whitePoint.X);
            double fy = LabFunction(xyz.Y / whitePoint.Y);
            double fz = LabFunction(xyz.Z / whitePoint.Z);
            lightness = Clamp(116D * fy - 16D, 0D, 100D);
            a = Clamp(500D * (fx - fy), -128D, 127D);
            b = Clamp(200D * (fy - fz), -128D, 127D);
        }

        private static double LabFunction(double value) => value > 216D / 24389D
            ? Math.Pow(Math.Max(0D, value), 1D / 3D)
            : (841D / 108D * value) + (4D / 29D);

        private static double Clamp(double value, double minimum, double maximum) =>
            value < minimum ? minimum : value > maximum ? maximum : value;
    }
}
