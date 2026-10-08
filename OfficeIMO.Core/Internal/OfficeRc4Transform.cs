using System;

namespace OfficeIMO.Core.Internal;

/// <summary>Shared stateful RC4 keystream primitive for qualified legacy carriers and fixed native header obfuscation.</summary>
internal sealed class OfficeRc4Transform {
    private readonly byte[] _state = new byte[256];
    private int _i;
    private int _j;

    internal OfficeRc4Transform(byte[] key) {
        if (key == null || key.Length == 0) throw new ArgumentException("A nonempty RC4 key is required.", nameof(key));
        for (int index = 0; index < _state.Length; index++) _state[index] = unchecked((byte)index);
        int j = 0;
        for (int index = 0; index < _state.Length; index++) {
            j = (j + _state[index] + key[index % key.Length]) & 0xFF;
            Swap(index, j);
        }
    }

    internal byte NextByte() {
        _i = (_i + 1) & 0xFF;
        _j = (_j + _state[_i]) & 0xFF;
        Swap(_i, _j);
        return _state[(_state[_i] + _state[_j]) & 0xFF];
    }

    private void Swap(int left, int right) {
        byte value = _state[left]; _state[left] = _state[right]; _state[right] = value;
    }
}
