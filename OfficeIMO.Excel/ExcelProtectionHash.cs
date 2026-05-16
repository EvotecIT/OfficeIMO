using System.Globalization;

namespace OfficeIMO.Excel {
    internal static class ExcelProtectionHash {
        // Excel's legacy worksheet/workbook protection hash salt.
        private const int LegacyPasswordSalt = 0xCE4B;

        internal static string? ResolveLegacyHash(string? password, string? legacyHash) {
            if (!string.IsNullOrWhiteSpace(legacyHash)) {
                return legacyHash!.Trim().ToUpperInvariant();
            }

            return string.IsNullOrEmpty(password) ? null : ComputeLegacyPasswordHash(password!);
        }

        internal static string ComputeLegacyPasswordHash(string password) {
            int hash = 0;
            for (int index = password.Length - 1; index >= 0; index--) {
                hash = Rotate(hash);
                hash ^= password[index];
            }

            hash = Rotate(hash);
            hash ^= password.Length;
            hash ^= LegacyPasswordSalt;
            return (hash & 0xFFFF).ToString("X4", CultureInfo.InvariantCulture);
        }

        private static int Rotate(int value) {
            return ((value >> 14) & 0x0001) | ((value << 1) & 0x7FFF);
        }
    }
}
