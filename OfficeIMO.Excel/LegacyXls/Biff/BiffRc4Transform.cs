namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal sealed class BiffRc4Transform {
        private readonly byte[] _state;
        private int _i;
        private int _j;

        internal BiffRc4Transform(byte[] key) {
            _state = new byte[256];
            for (int i = 0; i < _state.Length; i++) {
                _state[i] = (byte)i;
            }

            int j = 0;
            for (int i = 0; i < _state.Length; i++) {
                j = (j + _state[i] + key[i % key.Length]) & 0xff;
                Swap(i, j);
            }
        }

        internal byte NextByte() {
            _i = (_i + 1) & 0xff;
            _j = (_j + _state[_i]) & 0xff;
            Swap(_i, _j);
            return _state[(_state[_i] + _state[_j]) & 0xff];
        }

        internal static void Xor(byte[] key, byte[] data, int offset, int length) {
            var transform = new BiffRc4Transform(key);
            for (int i = 0; i < length; i++) {
                data[offset + i] = (byte)(data[offset + i] ^ transform.NextByte());
            }
        }

        private void Swap(int left, int right) {
            byte value = _state[left];
            _state[left] = _state[right];
            _state[right] = value;
        }
    }
}
