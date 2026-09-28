// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// OfficeIMO's copy is licensed under the repository's MIT license by the original author.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// VP8 arithmetic boolean reader with cooperative cancellation.
/// </summary>
internal sealed class OfficeVp8BoolDecoder
{
    private readonly byte[] _data;
    private readonly CancellationToken _cancellationToken;
    private int _reads;
    private int _offset;
    private int _range;
    private int _value;
    private int _count;
    private readonly bool _valid;
    private bool _paddingRead;
    private bool _exhausted;

    public OfficeVp8BoolDecoder(OfficeByteView data, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _cancellationToken = cancellationToken;
        _data = data.ToArray();
        _valid = _data.Length >= 2;

        if (!_valid)
        {
            return;
        }

        _range = 255;
        _value = (_data[0] << 8) | _data[1];
        _offset = 2;
        _count = 0;
    }

    public int BytesConsumed => _offset;

    public bool TryReadBool(int probability, out bool bit)
    {
        if ((_reads++ & 1023) == 0) _cancellationToken.ThrowIfCancellationRequested();
        bit = false;

        if (!_valid || _exhausted)
        {
            return false;
        }

        if ((uint)probability > 255)
        {
            return false;
        }

        // VP8 boolean arithmetic decoding step.
        var split = 1 + (((_range - 1) * probability) >> 8);
        var bigSplit = split << 8;

        if (_value >= bigSplit)
        {
            bit = true;
            _range -= split;
            _value -= bigSplit;
        }
        else
        {
            _range = split;
        }

        // Renormalize.
        while (_range < 128)
        {
            _range <<= 1;
            _value <<= 1;
            _count++;

            if (_count == 8)
            {
                _count = 0;
                _value |= ReadByte();
            }
        }

        // A second synthetic byte means this very symbol crossed the partition
        // boundary. Do not let a final header field or coefficient pass as valid.
        if (_exhausted)
        {
            bit = false;
            return false;
        }

        return true;
    }

    public bool TryReadLiteral(int bits, out int value)
    {
        value = 0;
        if (bits <= 0 || bits > 31)
        {
            return false;
        }

        var result = 0;
        for (var i = bits - 1; i >= 0; i--)
        {
            if (!TryReadBool(probability: 128, out var bit))
            {
                return false;
            }

            if (bit)
            {
                result |= 1 << i;
            }
        }

        value = result;
        return true;
    }

    public bool TryReadSignedLiteral(int bits, out int value)
    {
        value = 0;
        if (!TryReadLiteral(bits, out var magnitude))
        {
            return false;
        }

        if (!TryReadBool(probability: 128, out var signBit))
        {
            return false;
        }

        value = signBit ? -magnitude : magnitude;
        return true;
    }

    private int ReadByte()
    {
        if (_offset >= _data.Length)
        {
            // The arithmetic value includes one byte of lookahead. Allow one
            // zero byte to finish consuming the final real byte, but never let
            // an exhausted partition manufacture an unlimited stream of bits.
            if (_paddingRead) _exhausted = true;
            _paddingRead = true;
            return 0;
        }

        return _data[_offset++];
    }
}
