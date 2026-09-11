using System;

namespace OfficeIMO.Internal.Invoicing;

/// <summary>IBAN identity checks shared by native invoice mapping and PDF invoice inspection.</summary>
internal static class InvoiceBankAccountIdentity {
    internal static bool IsValidIban(string? value) {
        if (value == null || string.IsNullOrWhiteSpace(value)) return false;
        if (!TryNormalizeIban(value, out string? normalized)) {
            return false;
        }

        int modulo = 0;
        for (int i = 4; i < normalized!.Length; i++) {
            modulo = AppendIbanCharacterModulo(modulo, normalized[i]);
        }

        for (int i = 0; i < 4; i++) {
            modulo = AppendIbanCharacterModulo(modulo, normalized[i]);
        }

        return modulo == 1;
    }

    private static bool TryNormalizeIban(string value, out string? normalized) {
        var builder = new System.Text.StringBuilder(34);
        for (int i = 0; i < value.Length; i++) {
            char current = value[i];
            if (char.IsWhiteSpace(current)) {
                continue;
            }

            char upper = char.ToUpperInvariant(current);
            if (!((upper >= '0' && upper <= '9') || (upper >= 'A' && upper <= 'Z'))) {
                normalized = null;
                return false;
            }

            if (builder.Length == 34) { normalized = null; return false; }
            builder.Append(upper);
        }

        if (builder.Length < 15 || builder.Length > 34) {
            normalized = null;
            return false;
        }

        normalized = builder.ToString();
        if (!(normalized[0] >= 'A' && normalized[0] <= 'Z') ||
            !(normalized[1] >= 'A' && normalized[1] <= 'Z') ||
            !(normalized[2] >= '0' && normalized[2] <= '9') ||
            !(normalized[3] >= '0' && normalized[3] <= '9')) {
            normalized = null;
            return false;
        }

        return true;
    }

    private static int AppendIbanCharacterModulo(int modulo, char value) {
        if (value >= '0' && value <= '9') {
            return AppendIbanDigitModulo(modulo, value - '0');
        }

        int letterValue = value - 'A' + 10;
        modulo = AppendIbanDigitModulo(modulo, letterValue / 10);
        return AppendIbanDigitModulo(modulo, letterValue % 10);
    }

    private static int AppendIbanDigitModulo(int modulo, int digit) => (modulo * 10 + digit) % 97;
}
