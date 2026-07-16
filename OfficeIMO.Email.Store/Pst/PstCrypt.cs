namespace OfficeIMO.Email.Store;

internal static class PstCrypt {
    // Secondary permutation from [MS-PST] section 5.1 (mpbbS).
    private static readonly byte[] Secondary = {
        20, 83, 15, 86, 179, 200, 122, 156,
        235, 101, 72, 23, 22, 21, 159, 2,
        204, 84, 124, 131, 0, 13, 12, 11,
        162, 98, 168, 118, 219, 217, 237, 199,
        197, 164, 220, 172, 133, 116, 214, 208,
        167, 155, 174, 154, 150, 113, 102, 195,
        99, 153, 184, 221, 115, 146, 142, 132,
        125, 165, 94, 209, 93, 147, 177, 87,
        81, 80, 128, 137, 82, 148, 79, 78,
        10, 107, 188, 141, 127, 110, 71, 70,
        65, 64, 68, 1, 17, 203, 3, 63,
        247, 244, 225, 169, 143, 60, 58, 249,
        251, 240, 25, 48, 130, 9, 46, 201,
        157, 160, 134, 73, 238, 111, 77, 109,
        196, 45, 129, 52, 37, 135, 27, 136,
        170, 252, 6, 161, 18, 56, 253, 76,
        66, 114, 100, 19, 55, 36, 106, 117,
        119, 67, 255, 230, 180, 75, 54, 92,
        228, 216, 53, 61, 69, 185, 44, 236,
        183, 49, 43, 41, 7, 104, 163, 14,
        105, 123, 24, 158, 33, 57, 190, 40,
        26, 91, 120, 245, 35, 202, 42, 176,
        175, 62, 254, 4, 140, 231, 229, 152,
        50, 149, 211, 246, 74, 232, 166, 234,
        233, 243, 213, 47, 112, 32, 242, 31,
        5, 103, 173, 85, 16, 206, 205, 227,
        39, 59, 218, 186, 215, 194, 38, 212,
        145, 29, 210, 28, 34, 51, 248, 250,
        241, 90, 239, 207, 144, 182, 139, 181,
        189, 192, 191, 8, 151, 30, 108, 226,
        97, 224, 198, 193, 89, 171, 187, 88,
        222, 95, 223, 96, 121, 126, 178, 138
    };

    // Inverse permutation from [MS-PST] section 5.1 (mpbbI).
    private static readonly byte[] Inverse = {
        71, 241, 180, 230, 11, 106, 114, 72,
        133, 78, 158, 235, 226, 248, 148, 83,
        224, 187, 160, 2, 232, 90, 9, 171,
        219, 227, 186, 198, 124, 195, 16, 221,
        57, 5, 150, 48, 245, 55, 96, 130,
        140, 201, 19, 74, 107, 29, 243, 251,
        143, 38, 151, 202, 145, 23, 1, 196,
        50, 45, 110, 49, 149, 255, 217, 35,
        209, 0, 94, 121, 220, 68, 59, 26,
        40, 197, 97, 87, 32, 144, 61, 131,
        185, 67, 190, 103, 210, 70, 66, 118,
        192, 109, 91, 126, 178, 15, 22, 41,
        60, 169, 3, 84, 13, 218, 93, 223,
        246, 183, 199, 98, 205, 141, 6, 211,
        105, 92, 134, 214, 20, 247, 165, 102,
        117, 172, 177, 233, 69, 33, 112, 12,
        135, 159, 116, 164, 34, 76, 111, 191,
        31, 86, 170, 46, 179, 120, 51, 80,
        176, 163, 146, 188, 207, 25, 28, 167,
        99, 203, 30, 77, 62, 75, 27, 155,
        79, 231, 240, 238, 173, 58, 181, 89,
        4, 234, 64, 85, 37, 81, 229, 122,
        137, 56, 104, 82, 123, 252, 39, 174,
        215, 189, 250, 7, 244, 204, 142, 95,
        239, 53, 156, 132, 43, 21, 213, 119,
        52, 73, 182, 18, 10, 127, 113, 136,
        253, 157, 24, 65, 125, 147, 216, 88,
        44, 206, 254, 36, 175, 222, 184, 54,
        200, 161, 128, 166, 153, 152, 168, 47,
        14, 129, 101, 115, 228, 194, 162, 138,
        212, 225, 17, 208, 8, 139, 42, 242,
        237, 154, 100, 63, 193, 108, 249, 236
    };

    private static readonly byte[] Forward = CreateForwardPermutation();

    internal static void Decode(byte[] data, byte method, ulong bid = 0) {
        if (method == 0) return;
        if (method == 1) {
            for (int index = 0; index < data.Length; index++) data[index] = Inverse[data[index]];
            return;
        }
        if (method != 2) {
            throw new NotSupportedException(string.Concat("Unsupported PST/OST encryption method ",
                method.ToString(CultureInfo.InvariantCulture), "."));
        }

        ushort key = unchecked((ushort)((uint)bid ^ ((uint)bid >> 16)));
        for (int index = 0; index < data.Length; index++) {
            byte low = (byte)key;
            byte high = (byte)(key >> 8);
            byte value = unchecked((byte)(data[index] + low));
            value = Forward[value];
            value = unchecked((byte)(value + high));
            value = Secondary[value];
            value = unchecked((byte)(value - high));
            value = Inverse[value];
            data[index] = unchecked((byte)(value - low));
            key++;
        }
    }

    private static byte[] CreateForwardPermutation() {
        var result = new byte[Inverse.Length];
        for (int index = 0; index < Inverse.Length; index++) result[Inverse[index]] = checked((byte)index);
        return result;
    }
}
