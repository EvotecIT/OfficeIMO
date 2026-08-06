namespace OfficeIMO.Word;

/// <summary>
/// Specifies horizontal placement for text-box content.
/// </summary>
public enum WordTextBoxHorizontalAlignment {
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
/// Serializes text-box alignment values used by DrawingML.
/// </summary>
internal static class WordTextBoxHorizontalAlignmentSerializer {
    /// <summary>
    /// Convert alignment to string
    /// </summary>
    /// <param name="alignment"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentException"></exception>
    public static string ToString(WordTextBoxHorizontalAlignment alignment) {
        return alignment switch {
            WordTextBoxHorizontalAlignment.Left => "left",
            WordTextBoxHorizontalAlignment.Center => "center",
            WordTextBoxHorizontalAlignment.Right => "right",
            WordTextBoxHorizontalAlignment.Inside => "inside",
            WordTextBoxHorizontalAlignment.Outside => "outside",
            _ => throw new ArgumentException($"Invalid alignment value: {alignment}")
        };
    }

    /// <summary>
    /// Convert string to alignment
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public static WordTextBoxHorizontalAlignment FromString(string? value) {
        return value?.Trim().ToLowerInvariant() switch {
            "left" => WordTextBoxHorizontalAlignment.Left,
            "center" => WordTextBoxHorizontalAlignment.Center,
            "right" => WordTextBoxHorizontalAlignment.Right,
            "inside" => WordTextBoxHorizontalAlignment.Inside,
            "outside" => WordTextBoxHorizontalAlignment.Outside,
            _ => WordTextBoxHorizontalAlignment.Center
        };
    }
}
