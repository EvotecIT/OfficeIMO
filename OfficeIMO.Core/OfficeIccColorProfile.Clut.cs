using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeIccColorProfile {
    private sealed class MabClut {
        private readonly byte[] _payload;
        private readonly int[] _grid;
        private readonly int _outputChannels, _precision;
        internal MabClut(byte[] payload, int[] grid, int outputChannels, int precision) {
            _payload = payload; _grid = grid; _outputChannels = outputChannels; _precision = precision;
        }
        internal long RetainedByteCount => checked(88L + _payload.LongLength + _grid.LongLength * 4L);

        internal void Interpolate(double input0, double input1, double input2, double input3,
            out double output0, out double output1, out double output2, out double output3) =>
            Interpolate(new DeviceComponentValues(_grid.Length, input0, input1, input2, input3),
                out output0, out output1, out output2, out output3);

        internal void Interpolate(DeviceComponentValues input,
            out double output0, out double output1, out double output2, out double output3) {
            var positions = new DeviceComponentValues(_grid.Length,
                Clamp01(input[0]) * (_grid[0] - 1),
                Clamp01(input[1]) * (_grid[1] - 1),
                Clamp01(input[2]) * (_grid[2] - 1),
                _grid.Length > 3 ? Clamp01(input[3]) * (_grid[3] - 1) : 0D,
                _grid.Length > 4 ? Clamp01(input[4]) * (_grid[4] - 1) : 0D,
                _grid.Length > 5 ? Clamp01(input[5]) * (_grid[5] - 1) : 0D,
                _grid.Length > 6 ? Clamp01(input[6]) * (_grid[6] - 1) : 0D,
                _grid.Length > 7 ? Clamp01(input[7]) * (_grid[7] - 1) : 0D);
            output0 = output1 = output2 = output3 = 0D;
            int corners = 1 << _grid.Length;
            for (int corner = 0; corner < corners; corner++) {
                double weight = 1D;
                int gridIndex = 0;
                for (int channel = 0; channel < _grid.Length; channel++) {
                    bool upper = (corner & (1 << channel)) != 0;
                    int lower = Math.Min((int)positions[channel], _grid[channel] - 2);
                    double fraction = positions[channel] - lower;
                    weight *= upper ? fraction : 1D - fraction;
                    gridIndex = gridIndex * _grid[channel] + lower + (upper ? 1 : 0);
                }
                if (weight == 0D) continue;
                int offset = gridIndex * _outputChannels * _precision;
                output0 += ReadNormalized(offset) * weight;
                output1 += ReadNormalized(offset + _precision) * weight;
                output2 += ReadNormalized(offset + 2 * _precision) * weight;
                if (_outputChannels == 4) output3 += ReadNormalized(offset + 3 * _precision) * weight;
            }
        }
        private double ReadNormalized(int offset) => _precision == 1
            ? _payload[offset] / 255D : ReadUInt16(_payload, offset) / 65535D;
    }
}
