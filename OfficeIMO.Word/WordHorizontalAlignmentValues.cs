namespace OfficeIMO.Word;

/// <summary>
/// Word horizontal alignment values
/// </summary>
public enum WordHorizontalAlignmentValues {
    Left,
    Center,
    Right,
    Outside
}

/// <summary>
/// Class to help with horizontal alignment values
/// </summary>
public static class HorizontalAlignmentHelper {
    /// <summary>
    /// Convert alignment to string
    /// </summary>
    /// <param name="alignment"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentException"></exception>
    public static string ToString(WordHorizontalAlignmentValues alignment) {
        return alignment switch {
            WordHorizontalAlignmentValues.Left => "left",
            WordHorizontalAlignmentValues.Center => "center",
            WordHorizontalAlignmentValues.Right => "right",
            WordHorizontalAlignmentValues.Outside => "outside",
            _ => throw new ArgumentException($"Invalid alignment value: {alignment}")
        };
    }

    /// <summary>
    /// Convert string to alignment
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentException"></exception>
    public static WordHorizontalAlignmentValues FromString(string value) {
        return value.ToLowerInvariant() switch {
            "left" => WordHorizontalAlignmentValues.Left,
            "center" => WordHorizontalAlignmentValues.Center,
            "right" => WordHorizontalAlignmentValues.Right,
            "outside" => WordHorizontalAlignmentValues.Outside,
            _ => throw new ArgumentException($"Invalid alignment value: {value}")
        };
    }
}
