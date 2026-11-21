namespace OfficeIMO.Security {
    /// <summary>
    /// AES block cipher (FIPS 197) for hosts without a platform AES implementation, such as browser WebAssembly.
    /// </summary>
    /// <remarks>
    /// SubBytes and the key schedule look up a 256-byte S-box, so the cipher is not constant-time against cache-timing
    /// observers that share the processor. MixColumns and its inverse use branch-free arithmetic. Hosts with native AES
    /// should use the platform implementation.
    /// </remarks>
    internal sealed class ManagedAesBlockCipher : IDisposable {
        private const int BlockSize = 16;
        private static readonly byte[] SBox = new byte[256];
        private static readonly byte[] InverseSBox = new byte[256];

        private readonly bool _forEncryption;
        private readonly byte[] _roundKeys;
        private readonly int _rounds;
        private readonly byte[] _state = new byte[BlockSize];

        static ManagedAesBlockCipher() {
            // FIPS 197 section 5.1.1: the multiplicative inverse in GF(2^8), followed by the affine transformation.
            for (int value = 0; value < 256; value++) {
                byte inverse = value == 0 ? (byte)0 : Inverse((byte)value);
                int result = inverse ^ Rotate(inverse, 1) ^ Rotate(inverse, 2) ^ Rotate(inverse, 3) ^ Rotate(inverse, 4) ^ 0x63;
                SBox[value] = (byte)result;
                InverseSBox[result] = (byte)value;
            }
        }

        internal ManagedAesBlockCipher(byte[] key, bool forEncryption) {
            ArgumentNullException.ThrowIfNull(key);
            if (key.Length != 16 && key.Length != 24 && key.Length != 32) {
                throw new ArgumentException("AES keys must be 16, 24, or 32 bytes.", nameof(key));
            }

            _forEncryption = forEncryption;
            _rounds = key.Length / 4 + 6;
            _roundKeys = ExpandKey(key, _rounds);
        }

        internal void ProcessBlock(byte[] input, int inputOffset, byte[] output, int outputOffset) {
            ValidateBlock(input, inputOffset, nameof(input));
            ValidateBlock(output, outputOffset, nameof(output));

            Buffer.BlockCopy(input, inputOffset, _state, 0, BlockSize);
            if (_forEncryption) {
                Encrypt(_state);
            } else {
                Decrypt(_state);
            }

            Buffer.BlockCopy(_state, 0, output, outputOffset, BlockSize);
            Array.Clear(_state, 0, BlockSize);
        }

        public void Dispose() {
            Array.Clear(_roundKeys, 0, _roundKeys.Length);
            Array.Clear(_state, 0, _state.Length);
        }

        private void Encrypt(byte[] state) {
            AddRoundKey(state, 0);
            for (int round = 1; round < _rounds; round++) {
                SubBytes(state, SBox);
                ShiftRows(state);
                MixColumns(state);
                AddRoundKey(state, round);
            }

            SubBytes(state, SBox);
            ShiftRows(state);
            AddRoundKey(state, _rounds);
        }

        private void Decrypt(byte[] state) {
            // FIPS 197 section 5.3, InvCipher.
            AddRoundKey(state, _rounds);
            for (int round = _rounds - 1; round > 0; round--) {
                InverseShiftRows(state);
                SubBytes(state, InverseSBox);
                AddRoundKey(state, round);
                InverseMixColumns(state);
            }

            InverseShiftRows(state);
            SubBytes(state, InverseSBox);
            AddRoundKey(state, 0);
        }

        private void AddRoundKey(byte[] state, int round) {
            int offset = round * BlockSize;
            for (int index = 0; index < BlockSize; index++) {
                state[index] ^= _roundKeys[offset + index];
            }
        }

        private static void SubBytes(byte[] state, byte[] box) {
            for (int index = 0; index < BlockSize; index++) {
                state[index] = box[state[index]];
            }
        }

        // The state is column-major: byte r of column c is state[4 * c + r]. Row r rotates left by r columns.
        private static void ShiftRows(byte[] state) {
            byte value = state[1];
            state[1] = state[5]; state[5] = state[9]; state[9] = state[13]; state[13] = value;
            value = state[2]; state[2] = state[10]; state[10] = value;
            value = state[6]; state[6] = state[14]; state[14] = value;
            value = state[15];
            state[15] = state[11]; state[11] = state[7]; state[7] = state[3]; state[3] = value;
        }

        private static void InverseShiftRows(byte[] state) {
            byte value = state[13];
            state[13] = state[9]; state[9] = state[5]; state[5] = state[1]; state[1] = value;
            value = state[2]; state[2] = state[10]; state[10] = value;
            value = state[6]; state[6] = state[14]; state[14] = value;
            value = state[3];
            state[3] = state[7]; state[7] = state[11]; state[11] = state[15]; state[15] = value;
        }

        private static void MixColumns(byte[] state) {
            for (int column = 0; column < BlockSize; column += 4) {
                byte s0 = state[column], s1 = state[column + 1], s2 = state[column + 2], s3 = state[column + 3];
                byte all = (byte)(s0 ^ s1 ^ s2 ^ s3);
                // Each output byte is s_i ^ all ^ 2 * (s_i ^ s_(i+1)), the {02}{03}{01}{01} circulant.
                state[column] = (byte)(s0 ^ all ^ Double((byte)(s0 ^ s1)));
                state[column + 1] = (byte)(s1 ^ all ^ Double((byte)(s1 ^ s2)));
                state[column + 2] = (byte)(s2 ^ all ^ Double((byte)(s2 ^ s3)));
                state[column + 3] = (byte)(s3 ^ all ^ Double((byte)(s3 ^ s0)));
            }
        }

        private static void InverseMixColumns(byte[] state) {
            for (int column = 0; column < BlockSize; column += 4) {
                // {0e}{0b}{0d}{09} = {02}{03}{01}{01} applied after adding 4 * (s0 ^ s2) to s0 and s2 and 4 * (s1 ^ s3)
                // to s1 and s3, because the inverse circulant factors as {05}{00}{04}{00} times the forward one.
                byte even = Double(Double((byte)(state[column] ^ state[column + 2])));
                byte odd = Double(Double((byte)(state[column + 1] ^ state[column + 3])));
                state[column] ^= even;
                state[column + 1] ^= odd;
                state[column + 2] ^= even;
                state[column + 3] ^= odd;
            }

            MixColumns(state);
        }

        private static byte[] ExpandKey(byte[] key, int rounds) {
            // FIPS 197 section 5.2: Nb * (Nr + 1) words of four bytes.
            int keyWords = key.Length / 4;
            int totalWords = 4 * (rounds + 1);
            var schedule = new byte[totalWords * 4];
            Buffer.BlockCopy(key, 0, schedule, 0, key.Length);
            byte roundConstant = 1;
            var word = new byte[4];
            for (int index = keyWords; index < totalWords; index++) {
                Buffer.BlockCopy(schedule, (index - 1) * 4, word, 0, 4);
                if (index % keyWords == 0) {
                    byte first = word[0];
                    word[0] = (byte)(SBox[word[1]] ^ roundConstant);
                    word[1] = SBox[word[2]];
                    word[2] = SBox[word[3]];
                    word[3] = SBox[first];
                    roundConstant = Double(roundConstant);
                } else if (keyWords > 6 && index % keyWords == 4) {
                    for (int part = 0; part < 4; part++) {
                        word[part] = SBox[word[part]];
                    }
                }

                int previous = (index - keyWords) * 4;
                for (int part = 0; part < 4; part++) {
                    schedule[index * 4 + part] = (byte)(schedule[previous + part] ^ word[part]);
                }
            }

            Array.Clear(word, 0, word.Length);
            return schedule;
        }

        // Multiplication by {02} modulo x^8 + x^4 + x^3 + x + 1, without a data-dependent branch.
        private static byte Double(byte value) => (byte)((value << 1) ^ (0x1B & -(value >> 7)));

        private static byte Multiply(byte left, byte right) {
            byte product = 0;
            for (int bit = 0; bit < 8; bit++) {
                product ^= (byte)(left & -(right & 1));
                left = Double(left);
                right >>= 1;
            }

            return product;
        }

        // a^254 is the inverse of a non-zero element of GF(2^8).
        private static byte Inverse(byte value) {
            byte result = 1;
            byte power = value;
            for (int exponent = 254; exponent > 0; exponent >>= 1) {
                if ((exponent & 1) != 0) {
                    result = Multiply(result, power);
                }

                power = Multiply(power, power);
            }

            return result;
        }

        private static int Rotate(byte value, int shift) => ((value << shift) | (value >> (8 - shift))) & 0xFF;

        private static void ValidateBlock(byte[] buffer, int offset, string parameterName) {
            ArgumentNullException.ThrowIfNull(buffer, parameterName);

            if (offset < 0 || offset > buffer.Length - BlockSize) {
                throw new ArgumentOutOfRangeException(parameterName, "A complete 16-byte AES block is required.");
            }
        }
    }
}
