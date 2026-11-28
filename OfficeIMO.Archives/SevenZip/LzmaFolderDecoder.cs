namespace OfficeIMO.Archives.SevenZip;

/// <summary>Decodes the LZMA or LZMA2 stream of one 7z folder into its admitted output buffer.</summary>
internal sealed class LzmaFolderDecoder {
    private const int PositionContextCount = 16;
    private const int LiteralModelSize = 768;
    private readonly byte[] _output;
    private readonly uint _dictionarySize;
    private readonly ushort[] _matches = new ushort[12 * PositionContextCount];
    private readonly ushort[] _shortRepeats = new ushort[12 * PositionContextCount];
    private readonly ushort[] _repeatKinds = new ushort[12 * 4];
    private readonly ushort[] _distanceSlots = new ushort[4 * 64];
    private readonly ushort[] _smallDistances = new ushort[115];
    private readonly ushort[] _distanceAlignment = new ushort[16];
    private readonly LengthModel _newLength = new();
    private readonly LengthModel _repeatLength = new();
    private readonly uint[] _recent = new uint[4];
    private ushort[] _literals = Array.Empty<ushort>();
    private int _position;
    private int _historyStart;
    private int _state;
    private int _literalContextBits;
    private int _literalPositionMask;
    private int _positionMask;

    private LzmaFolderDecoder(byte[] output, uint dictionarySize) {
        _output = output;
        _dictionarySize = dictionarySize;
    }

    internal static void DecodeLzma(Stream source, byte[] properties, byte[] output, ArchiveBlockBudget budget, CancellationToken token) {
        uint dictionary = SevenZipInput.UInt32At(properties, 1);
        var decoder = new LzmaFolderDecoder(output, Math.Max(4096u, dictionary));
        decoder.ConfigureModels(properties[0], lzma2: false, budget);
        decoder.DecodeSegment(source, source.Length, output.Length, allowEndMarker: true, token);
    }

    internal static void DecodeLzma2(Stream source, byte[] properties, byte[] output, ArchiveBlockBudget budget, CancellationToken token) {
        int property = properties[0];
        uint dictionary = property == 40 ? uint.MaxValue : (uint)(2 + property % 2) << (11 + property / 2);
        var decoder = new LzmaFolderDecoder(output, dictionary);
        bool requireDictionary = true;
        bool requireProperties = true;
        while (true) {
            token.ThrowIfCancellationRequested();
            int control = Octet(source);
            if (control == 0) break;
            if (control == 1 || control >= 0xe0) {
                decoder._historyStart = decoder._position;
                requireDictionary = false;
                requireProperties = true;
            }
            if (requireDictionary) throw new InvalidDataException("The 7z LZMA2 stream is missing its initial dictionary reset.");

            if (control is 1 or 2) {
                int plainSize = ReadChunkSize(source);
                decoder.EnsureFits(plainSize);
                SevenZipInput.Fill(source, decoder._output, decoder._position, plainSize, token);
                // An uncompressed chunk updates the history without rebuilding the range models.
                decoder._position += plainSize;
                continue;
            }
            if (control < 0x80) throw new InvalidDataException("The 7z LZMA2 chunk control byte is reserved.");

            int plainLength = ((control & 31) << 16) + ReadChunkSize(source);
            int packedLength = ReadChunkSize(source);
            bool newProperties = control >= 0xc0;
            bool resetModels = control >= 0xa0;
            if (requireProperties && !newProperties)
                throw new InvalidDataException("The 7z LZMA2 chunk is missing new properties after a dictionary reset.");
            if (resetModels) {
                // Many short runs that reset the probability tables cost more than their output size.
                if (newProperties) decoder.ConfigureModels(Octet(source), lzma2: true, budget);
                else decoder.ResetModels(budget);
                requireProperties = false;
            }
            decoder.DecodeSegment(source, packedLength, plainLength, allowEndMarker: false, token);
        }
        if (decoder._position != output.Length || source.Position != source.Length)
            throw new InvalidDataException("The 7z LZMA2 stream does not end at its declared boundaries.");
    }

    private void ConfigureModels(int properties, bool lzma2, ArchiveBlockBudget budget) {
        int literalBits = properties % 9;
        int positionBits = properties / 9 % 5;
        int matchPositionBits = properties / 45;
        if (matchPositionBits > 4 || lzma2 && literalBits + positionBits > 4)
            throw new InvalidDataException("The 7z LZMA literal and position properties are invalid.");
        _literalContextBits = literalBits;
        _literalPositionMask = (1 << positionBits) - 1;
        _positionMask = (1 << matchPositionBits) - 1;
        int count = LiteralModelSize << (literalBits + positionBits);
        ChargeReset(budget, count);
        if (_literals.Length != count) _literals = new ushort[count];
        ResetModels();
    }

    private void ResetModels(ArchiveBlockBudget budget) {
        ChargeReset(budget, _literals.Length);
        ResetModels();
    }

    private static void ChargeReset(ArchiveBlockBudget budget, int literals) =>
        budget.Charge(0, 4096 + 2L * literals);

    private void ResetModels() {
        _state = 0;
        Array.Clear(_recent, 0, _recent.Length);
        Center(_matches);
        Center(_shortRepeats);
        Center(_repeatKinds);
        Center(_distanceSlots);
        Center(_smallDistances);
        Center(_distanceAlignment);
        Center(_literals);
        _newLength.Reset();
        _repeatLength.Reset();
    }

    private void DecodeSegment(Stream source, long packedBytes, int decodedBytes, bool allowEndMarker, CancellationToken token) {
        EnsureFits(decodedBytes);
        var reader = new LzmaRangeDecoder(source, packedBytes);
        int end = _position + decodedBytes;
        int nextCheck = _position;
        while (_position < end) {
            if (_position >= nextCheck) {
                token.ThrowIfCancellationRequested();
                nextCheck = _position + Math.Min(8192, end - _position);
            }
            if (!DecodeSymbol(reader, end)) throw new InvalidDataException("The LZMA stream ended before its declared output size.");
        }
        if (!reader.Finished && (!allowEndMarker || DecodeSymbol(reader, end) || !reader.Finished))
            throw new InvalidDataException("The LZMA stream does not finish at its declared packed size.");
    }

    private bool DecodeSymbol(LzmaRangeDecoder reader, int end) {
        int context = (_position - _historyStart) & _positionMask;
        int stateAndContext = _state * PositionContextCount + context;
        if (reader.Bit(_matches, stateAndContext) == 0) {
            if (_position == end) throw new InvalidDataException("The LZMA stream exceeds its declared output size.");
            byte literal = ReadLiteral(reader);
            _output[_position++] = literal;
            _state = _state < 4 ? 0 : _state - (_state < 10 ? 3 : 6);
            return true;
        }

        int length = 1;
        int kind = _state * 4;
        if (reader.Bit(_repeatKinds, kind) == 0) {
            length = _newLength.Read(reader, context);
            uint distance = ReadDistance(reader, length - 2);
            if (distance == uint.MaxValue) return false;
            Remember(3, distance);
            _state = _state < 7 ? 7 : 10;
        } else {
            bool shortRepeat = reader.Bit(_repeatKinds, kind + 1) == 0;
            if (shortRepeat) {
                shortRepeat = reader.Bit(_shortRepeats, stateAndContext) == 0;
            } else {
                int index = reader.Bit(_repeatKinds, kind + 2) == 0 ? 1 : 2 + reader.Bit(_repeatKinds, kind + 3);
                Remember(index, _recent[index]);
            }
            if (!shortRepeat) length = _repeatLength.Read(reader, context);
            _state = _state < 7 ? (shortRepeat ? 9 : 8) : 11;
        }
        if (length > end - _position) throw new InvalidDataException("An LZMA match exceeds its declared output size.");
        int from = HistoryOffset(_recent[0]);
        for (int i = 0; i < length; i++) _output[_position++] = _output[from++];
        return true;
    }

    private byte ReadLiteral(LzmaRangeDecoder reader) {
        int previous = _position > _historyStart ? _output[_position - 1] : 0;
        int positionContext = (_position - _historyStart) & _literalPositionMask;
        int context = (positionContext << _literalContextBits) | (previous >> (8 - _literalContextBits));
        int baseIndex = context * LiteralModelSize;
        int prefix = 1;
        if (_state >= 7) {
            int neighbour = _output[HistoryOffset(_recent[0])];
            for (int shift = 7; shift >= 0; shift--) {
                int expected = (neighbour >> shift) & 1;
                int bit = reader.Bit(_literals, baseIndex + 256 * (expected + 1) + prefix);
                prefix = 2 * prefix + bit;
                if (bit != expected) break;
            }
        }
        while (prefix < 256) prefix = 2 * prefix + reader.Bit(_literals, baseIndex + prefix);
        return (byte)(prefix - 256);
    }

    private uint ReadDistance(LzmaRangeDecoder reader, int length) {
        int slot = reader.Tree(_distanceSlots, Math.Min(3, length) * 64, 6);
        if (slot < 4) return (uint)slot;
        int bits = slot / 2 - 1;
        uint start = (uint)(2 + slot % 2) << bits;
        uint suffix = slot < 14
            ? reader.ReversedTree(_smallDistances, (int)start - slot, bits)
            : (reader.Bypass(bits - 4) << 4) | reader.ReversedTree(_distanceAlignment, 0, 4);
        return start | suffix;
    }

    private void Remember(int last, uint distance) {
        for (int i = last; i > 0; i--) _recent[i] = _recent[i - 1];
        _recent[0] = distance;
    }

    private int HistoryOffset(uint distance) {
        if (distance >= _dictionarySize || distance >= (uint)(_position - _historyStart))
            throw new InvalidDataException("An LZMA match points outside its initialized dictionary.");
        return _position - (int)distance - 1;
    }

    private void EnsureFits(int count) {
        if (count > _output.Length - _position) throw new InvalidDataException("A 7z compressed chunk exceeds its declared output size.");
    }

    private static int Octet(Stream stream) {
        int value = stream.ReadByte();
        if (value < 0) throw new InvalidDataException("The 7z LZMA2 stream is truncated.");
        return value;
    }

    private static int ReadChunkSize(Stream stream) => (Octet(stream) << 8) + Octet(stream) + 1;

    private static void Center(ushort[] weights) {
        for (int i = 0; i < weights.Length; i++) weights[i] = 1024;
    }

    private sealed class LengthModel {
        private readonly ushort[] _tiers = new ushort[2];
        private readonly ushort[] _positions = new ushort[2 * PositionContextCount * 8];
        private readonly ushort[] _long = new ushort[256];

        internal void Reset() {
            Center(_tiers);
            Center(_positions);
            Center(_long);
        }

        internal int Read(LzmaRangeDecoder reader, int context) {
            if (reader.Bit(_tiers, 0) == 0) return 2 + reader.Tree(_positions, context * 8, 3);
            if (reader.Bit(_tiers, 1) == 0) return 10 + reader.Tree(_positions, (PositionContextCount + context) * 8, 3);
            return 18 + reader.Tree(_long, 0, 8);
        }
    }
}
