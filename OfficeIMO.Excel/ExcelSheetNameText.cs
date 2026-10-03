namespace OfficeIMO.Excel;

/// <summary>UTF-16 length bounds shared by worksheet-name normalization and collision handling.</summary>
internal static class ExcelSheetNameText {
    internal static string Truncate(string value, int maximumLength) {
        if (maximumLength < 0) throw new System.ArgumentOutOfRangeException(nameof(maximumLength));
        if (value.Length <= maximumLength) return value;
        int length = maximumLength;
        if (length > 0 && char.IsHighSurrogate(value[length - 1]) && char.IsLowSurrogate(value[length])) {
            length--;
        }
        return value.Substring(0, length);
    }
}
