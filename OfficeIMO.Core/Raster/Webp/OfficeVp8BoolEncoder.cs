// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// VP8 boolean encoder for managed VP8 bitstreams.
/// </summary>
internal sealed class OfficeVp8BoolEncoder {
    private static readonly byte[] Norm = {
        7, 6, 6, 5, 5, 5, 5, 4, 4, 4, 4, 4, 4, 4, 4,
        3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3,
        2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2,
        2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2,
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
        0,
    };

    private static readonly byte[] NewRange = {
        127, 127, 191, 127, 159, 191, 223, 127, 143, 159, 175, 191, 207, 223, 239,
        127, 135, 143, 151, 159, 167, 175, 183, 191, 199, 207, 215, 223, 231, 239,
        247, 127, 131, 135, 139, 143, 147, 151, 155, 159, 163, 167, 171, 175, 179,
        183, 187, 191, 195, 199, 203, 207, 211, 215, 219, 223, 227, 231, 235, 239,
        243, 247, 251, 127, 129, 131, 133, 135, 137, 139, 141, 143, 145, 147, 149,
        151, 153, 155, 157, 159, 161, 163, 165, 167, 169, 171, 173, 175, 177, 179,
        181, 183, 185, 187, 189, 191, 193, 195, 197, 199, 201, 203, 205, 207, 209,
        211, 213, 215, 217, 219, 221, 223, 225, 227, 229, 231, 233, 235, 237, 239,
        241, 243, 245, 247, 249, 251, 253, 127,
    };

    private readonly int _maximumBytes;
    private readonly CancellationToken _cancellationToken;
    private readonly OfficeVp8EncodingMemory _memory;
    private int _range;
    private int _value;
    private int _run;
    private int _nbBits;
    private byte[] _buffer;
    private int _pos;

    internal OfficeVp8BoolEncoder(int expectedSize, int maximumBytes, CancellationToken cancellationToken, OfficeVp8EncodingMemory memory) {
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumBytes = maximumBytes;
        _cancellationToken = cancellationToken;
        _memory = memory;
        expectedSize = Math.Max(1, Math.Min(expectedSize, maximumBytes));
        _memory.Reserve(expectedSize + 24L);
        _range = 255 - 1;
        _value = 0;
        _run = 0;
        _nbBits = -8;
        _buffer = new byte[expectedSize];
        _pos = 0;
    }

    public void WriteBool(int probability, bool bit) {
        if ((uint)probability > 255) {
            throw new ArgumentOutOfRangeException(nameof(probability));
        }

        var split = (_range * probability) >> 8;
        if (bit) {
            _value += split + 1;
            _range -= split + 1;
        } else {
            _range = split;
        }

        if (_range < 127) {
            var shift = Norm[_range];
            _range = NewRange[_range];
            _value <<= shift;
            _nbBits += shift;
            if (_nbBits > 0) {
                Flush();
            }
        }
    }

    public void WriteLiteral(int value, int bits) {
        if (bits <= 0 || bits >= 32) throw new ArgumentOutOfRangeException(nameof(bits));
        var mask = 1 << (bits - 1);
        while (mask != 0) {
            WriteBoolUniform((value & mask) != 0);
            mask >>= 1;
        }
    }

    public byte[] Finish() {
        WriteLiteral(0, 9 - _nbBits);
        _nbBits = 0;
        Flush();
        if (_pos == _buffer.Length) return _buffer;
        _memory.Reserve(_pos + 24L);
        var output = new byte[_pos];
        Buffer.BlockCopy(_buffer, 0, output, 0, _pos);
        return output;
    }

    private void WriteBoolUniform(bool bit) {
        var split = _range >> 1;
        if (bit) {
            _value += split + 1;
            _range -= split + 1;
        } else {
            _range = split;
        }

        if (_range < 127) {
            _range = NewRange[_range];
            _value <<= 1;
            _nbBits += 1;
            if (_nbBits > 0) {
                Flush();
            }
        }
    }

    private void Flush() {
        var shift = 8 + _nbBits;
        var bits = _value >> shift;
        if (_nbBits < 0) return;

        _value -= bits << shift;
        _nbBits -= 8;

        if ((bits & 0xFF) != 0xFF) {
            var pos = _pos;
            EnsureCapacity(_run + 1);
            if ((bits & 0x100) != 0) {
                if (pos > 0) {
                    _buffer[pos - 1]++;
                }
            }

            if (_run > 0) {
                var value = (bits & 0x100) != 0 ? (byte)0x00 : (byte)0xFF;
                while (_run > 0) {
                    _buffer[pos++] = value;
                    _run--;
                }
            }

            _buffer[pos++] = (byte)(bits & 0xFF);
            _pos = pos;
        } else {
            _run++;
        }
    }

    private void EnsureCapacity(int extraBytes) {
        if (extraBytes <= 0) return;
        _cancellationToken.ThrowIfCancellationRequested();
        long requested = (long)_pos + extraBytes;
        if (requested > _maximumBytes) throw new ArgumentException("VP8 output exceeds encoded-size limits.");
        int needed = (int)requested;
        if (needed <= _buffer.Length) return;

        var newSize = _buffer.Length * 2;
        if (newSize < needed) newSize = needed;
        if (newSize < 1024) newSize = 1024;
        newSize = Math.Min(newSize, _maximumBytes);
        _memory.ReplaceBuffer(_buffer.Length, newSize);
        Array.Resize(ref _buffer, newSize);
    }
}
