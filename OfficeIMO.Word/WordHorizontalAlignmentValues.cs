namespace OfficeIMO.Word;

/// <summary>
/// Word horizontal alignment values
/// </summary>
public enum WordHorizontalAlignmentValues {
    /// <summary>
    /// Content is aligned to the left.
    /// </summary>
    Left = 0,
    /// <summary>
    /// Content is centered.
    /// </summary>
    Center = 1,
    /// <summary>
    /// Content is aligned to the right.
    /// </summary>
    Right = 2,
    /// <summary>
    /// Content is aligned to the outside of odd or even pages.
    /// </summary>
    Outside = 3,
    /// <summary>
    /// Content is aligned to the inside of odd or even pages.
    /// </summary>
    Inside = 4
}

/// <summary>
/// Class to help with horizontal alignment values
/// </summary>
internal static class HorizontalAlignmentHelper {
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
            WordHorizontalAlignmentValues.Inside => "inside",
            WordHorizontalAlignmentValues.Outside => "outside",
            _ => throw new ArgumentException($"Invalid alignment value: {alignment}")
        };
    }

    /// <summary>
    /// Convert string to alignment
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public static WordHorizontalAlignmentValues FromString(string? value) {
        return value?.Trim().ToLowerInvariant() switch {
            "left" => WordHorizontalAlignmentValues.Left,
            "center" => WordHorizontalAlignmentValues.Center,
            "right" => WordHorizontalAlignmentValues.Right,
            "inside" => WordHorizontalAlignmentValues.Inside,
            "outside" => WordHorizontalAlignmentValues.Outside,
            _ => WordHorizontalAlignmentValues.Center
        };
    }
}
