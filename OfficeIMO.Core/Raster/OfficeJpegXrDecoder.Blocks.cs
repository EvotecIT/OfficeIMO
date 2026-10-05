using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    private sealed class AdaptiveVlc {
        private readonly Codebook[] _codes;
        private readonly int[][] _deltas;
        private int _table = 1, _low, _high, _lowDelta, _highDelta = 1;
        internal AdaptiveVlc(Codebook[] codes, int[][] deltas) { _codes = codes; _deltas = deltas; }

        internal int Read(Bits bits) {
            int value = _codes[_table].Read(bits);
            _low += _deltas[_lowDelta][value]; _high += _deltas[_highDelta][value];
            return value;
        }

        internal void Adapt() {
            int previous = _table;
            if (_low < -8 && _table != 0) _table--;
            else if (_high > 8 && _table != _codes.Length - 1) _table++;
            if (previous != _table) {
                _low = _high = 0;
                _lowDelta = Math.Max(0, _table - 1);
                _highDelta = Math.Min(_table, _codes.Length - 2);
            } else {
                _low = Math.Max(-64, Math.Min(64, _low)); _high = Math.Max(-64, Math.Min(64, _high));
            }
        }
    }

    private sealed class BlockContext {
        private readonly AdaptiveVlc[] _first = {
            new(FirstCodes, FirstDeltas), new(FirstCodes, FirstDeltas)
        };
        private readonly AdaptiveVlc[,] _index = {
            { new(IndexCodes, IndexDeltas), new(IndexCodes, IndexDeltas) },
            { new(IndexCodes, IndexDeltas), new(IndexCodes, IndexDeltas) }
        };
        private readonly BinaryVlc[] _absolute = { new(), new() };
        internal readonly int[] Runs = new int[15], Levels = new int[15];

        // T.832 8.7.18.5. Run/level lists contain at most fifteen AC coefficients.
        internal int Read(Bits bits, bool chroma) {
            int channel = chroma ? 1 : 0, first = _first[channel].Read(bits);
            bool negative = bits.Flag();
            int preceding = first & 1, following = first >> 2, context = preceding & following;
            Levels[0] = (first & 2) != 0 ? _absolute[context].ReadAbsolute(bits) : 1;
            if (negative) Levels[0] = -Levels[0];
            Runs[0] = preceding == 0 ? ReadRun(bits, 14) : 0;
            int location = Runs[0] + 2, count = 1;
            while (following != 0) {
                if (count == 15 || location >= 16) throw new FormatException("JPEG-XR AC run exceeds its block.");
                Runs[count] = (following & 1) == 0 ? ReadRun(bits, 15 - location) : 0;
                location += Runs[count] + 1;
                if (location > 16) throw new FormatException("JPEG-XR AC run exceeds its block.");
                int index = location < 15 ? _index[channel, context].Read(bits)
                    : location == 15 ? FinalIndexCodes.Read(bits) : (int)bits.Read(1);
                following = index >> 1; context &= following;
                negative = bits.Flag();
                Levels[count] = (index & 1) != 0 ? _absolute[context].ReadAbsolute(bits) : 1;
                if (negative) Levels[count] = -Levels[count];
                count++;
            }
            return count;
        }

        internal void Adapt() {
            foreach (AdaptiveVlc vlc in _first) vlc.Adapt();
            foreach (AdaptiveVlc vlc in _index) vlc.Adapt();
            foreach (BinaryVlc vlc in _absolute) vlc.Adapt();
        }

        private static int ReadRun(Bits bits, int maximum) {
            if (maximum < 1 || maximum > 14) throw new FormatException("JPEG-XR run range is invalid.");
            int run;
            if (maximum < 5) {
                run = 1;
                while (run < maximum && !bits.Flag()) run++;
            } else {
                int index = RunCodes.Read(bits) + 5 * RunBin[maximum];
                run = RunBase[index] + (int)bits.Read(RunBits[index]);
            }
            if (run > maximum) throw new FormatException("JPEG-XR run exceeds the available coefficients.");
            return run;
        }
    }
}
