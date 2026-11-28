namespace OfficeIMO.Entities;

/// <summary>An offset and shape within the recognition context's text.</summary>
public readonly struct EntityToken {
    internal EntityToken(EntityTokenShape shape, int start, int length, int characterCount, bool touchesWord) {
        Shape = shape;
        Start = start;
        Length = length;
        CharacterCount = characterCount;
        TouchesWord = touchesWord;
    }

    /// <summary>The shape the tokenizer recognized.</summary>
    public EntityTokenShape Shape { get; }
    /// <summary>UTF-16 offset relative to the current segment.</summary>
    public int Start { get; }
    /// <summary>Written length, including any internal separators.</summary>
    public int Length { get; }
    /// <summary>Number of characters excluding internal separators.</summary>
    public int CharacterCount { get; }
    /// <summary>Whether a letter or digit touches either end of the token.</summary>
    public bool TouchesWord { get; }
}
