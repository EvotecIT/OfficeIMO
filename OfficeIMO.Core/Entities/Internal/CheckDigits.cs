namespace OfficeIMO.Entities.Internal;

internal static class CheckDigits {
    internal static int WeightedSum(string value, string weights, int start = 0, bool expandProducts = false) {
        int sum = 0;
        for (int index = 0; index < weights.Length; index++) {
            int product = AlphaNumericValue(value[start + index]) * (weights[index] - '0');
            sum += expandProducts ? product / 10 + product % 10 : product;
        }
        return sum;
    }

    internal static int AlphaNumericValue(char value) => value <= '9' ? value - '0' : value - 'A' + 10;

    internal static int AppendMod97(int remainder, char value) {
        int number = AlphaNumericValue(value);
        return (remainder * (number < 10 ? 10 : 100) + number) % 97;
    }

    internal static int Mod97(string value, int start = 0, int rotate = 0) {
        int remainder = 0;
        for (int index = start + rotate; index < value.Length; index++) remainder = AppendMod97(remainder, value[index]);
        for (int index = start; index < start + rotate; index++) remainder = AppendMod97(remainder, value[index]);
        return remainder;
    }

    internal static bool HybridMod11_10(string value) {
        int product = 10;
        for (int index = 0; index < value.Length - 1; index++) {
            int sum = (product + value[index] - '0') % 10;
            product = 2 * (sum == 0 ? 10 : sum) % 11;
        }
        return (11 - product) % 10 == value[value.Length - 1] - '0';
    }

    internal static bool Luhn(string value, int? length = null) {
        int sum = 0;
        bool doubleDigit = false;
        for (int index = (length ?? value.Length) - 1; index >= 0; index--) {
            int number = AlphaNumericValue(value[index]);
            sum += LuhnDigit(number % 10, doubleDigit);
            doubleDigit = !doubleDigit;
            if (number >= 10) {
                sum += LuhnDigit(number / 10, doubleDigit);
                doubleDigit = !doubleDigit;
            }
        }
        return sum % 10 == 0;
    }

    private static int LuhnDigit(int digit, bool doubleDigit) => doubleDigit ? digit * 2 - (digit >= 5 ? 9 : 0) : digit;
}
