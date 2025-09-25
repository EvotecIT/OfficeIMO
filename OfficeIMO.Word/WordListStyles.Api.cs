using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Provides access to the built-in list styles and helper methods for
/// working with numbering definitions.
/// </summary>
public static partial class WordListStyles {
    /// <summary>
    /// Retrieves the abstract numbering definition for the specified built-in list style.
    /// </summary>
    /// <param name="style">The list style to retrieve.</param>
    /// <returns>The corresponding <see cref="AbstractNum"/> definition.</returns>
    public static AbstractNum GetStyle(WordListStyle style) {
        switch (style) {
            case WordListStyle.Bulleted: return Bulleted;
            case WordListStyle.ArticleSections: return ArticleSections;
            case WordListStyle.Headings111: return Headings111;
            case WordListStyle.HeadingIA1: return HeadingIA1;
            case WordListStyle.Chapters: return Chapters;
            case WordListStyle.BulletedChars: return BulletedChars;
            case WordListStyle.Heading1ai: return Heading1ai;
            case WordListStyle.Headings111Shifted: return Headings111Shifted;
            case WordListStyle.LowerLetterWithBracket: return LowerLetterWithBracket;
            case WordListStyle.LowerLetterWithDot: return LowerLetterWithDot;
            case WordListStyle.UpperLetterWithDot: return UpperLetterWithDot;
            case WordListStyle.UpperLetterWithBracket: return UpperLetterWithBracket;
            case WordListStyle.Custom: return Custom;
            case WordListStyle.Numbered: return Numbered;
        }
        throw new ArgumentOutOfRangeException(nameof(style));
    }

    /// <summary>
    /// Attempts to match an <see cref="AbstractNum"/> to one of the built-in list styles.
    /// </summary>
    /// <param name="abstractNum">The abstract numbering definition to compare.</param>
    /// <returns>The matching <see cref="WordListStyle"/> or <see cref="WordListStyle.Custom"/> when no match is found.</returns>
    public static WordListStyle MatchStyle(AbstractNum abstractNum) {
        if (abstractNum == null) throw new ArgumentNullException(nameof(abstractNum));

        var templateCode = abstractNum.GetFirstChild<TemplateCode>()?.Val?.Value;
        return templateCode switch {
            "934E79A6" => WordListStyle.Bulleted,
            "04090023" => WordListStyle.ArticleSections,
            "04090025" => WordListStyle.Headings111,
            "04090027" => WordListStyle.HeadingIA1,
            "04090029" => WordListStyle.Chapters,
            "04090021" => WordListStyle.BulletedChars,
            "0409001D" => WordListStyle.Heading1ai,
            "0409001F" => WordListStyle.Headings111Shifted,
            "BB9E481E" => WordListStyle.LowerLetterWithBracket,
            "73ECA528" => WordListStyle.LowerLetterWithDot,
            "76643E8A" => WordListStyle.UpperLetterWithDot,
            "76643E8C" => WordListStyle.UpperLetterWithBracket,
            "25BE4216" => WordListStyle.Numbered,
            _ => WordListStyle.Custom
        };
    }

    /// <summary>
    /// Creates numbering definitions for basic bulleted and numbered lists.
    /// </summary>
    /// <param name="document">The document to base the numbering definitions on.</param>
    /// <param name="bulletNumberId">Returns the numbering identifier for bullets.</param>
    /// <param name="orderedNumberId">Returns the numbering identifier for ordered lists.</param>
    /// <returns>A <see cref="Numbering"/> instance containing the definitions.</returns>
    public static Numbering CreateDefaultNumberingDefinitions(WordprocessingDocument document, out int bulletNumberId, out int orderedNumberId) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        InitializeAbstractNumberId(document);

        AbstractNum bulletAbstract = GetStyle(WordListStyle.Bulleted);
        AbstractNum orderedAbstract = GetStyle(WordListStyle.Numbered);

        var bulletAbstractNumberId = bulletAbstract.AbstractNumberId?.Value
            ?? throw new InvalidOperationException("Bulleted style is missing an abstract number identifier.");
        var orderedAbstractNumberId = orderedAbstract.AbstractNumberId?.Value
            ?? throw new InvalidOperationException("Numbered style is missing an abstract number identifier.");

        bulletNumberId = 1;
        orderedNumberId = 2;

        NumberingInstance bulletInstance = new NumberingInstance(new AbstractNumId { Val = bulletAbstractNumberId }) { NumberID = bulletNumberId };
        NumberingInstance orderedInstance = new NumberingInstance(new AbstractNumId { Val = orderedAbstractNumberId }) { NumberID = orderedNumberId };

        return new Numbering(bulletAbstract, bulletInstance, orderedAbstract, orderedInstance);
    }

    /// <summary>
    /// The next abstract number identifier stored to be used when creating new abstract numbers
    /// </summary>
    private static int nextAbstractNumberId;

    /// <summary>
    /// Generates a unique NSID value using a GUID
    /// </summary>
    /// <returns></returns>
    private static string GenerateNsidValue() {
        return Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant();
    }

    /// <summary>
    /// Initializes the abstract number identifier, starting from the highest value currently in use
    /// It makes sure that the next abstract number identifier is unique
    /// </summary>
    /// <param name="document">The document.</param>
    internal static void InitializeAbstractNumberId(WordprocessingDocument document) {
        // Find the highest AbstractNumberId currently in use

        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var mainPart = document.MainDocumentPart;
        if (mainPart == null) {
            nextAbstractNumberId = 0;
            return;
        }

        var numberingDefinitionPart = mainPart.NumberingDefinitionsPart;
        if (numberingDefinitionPart?.Numbering == null) {
            // No numbering definitions part found, so no abstract numbers are in use
            nextAbstractNumberId = 0;
            return;
        }

        nextAbstractNumberId = numberingDefinitionPart
            .Numbering
            .Descendants<AbstractNum>()
            .Select(an => an.AbstractNumberId?.Value ?? -1)
            .DefaultIfEmpty(-1)
            .Max();

        // Start assigning AbstractNumberId values from the next number
        nextAbstractNumberId++;
    }

    /// <summary>
    /// Creates the new abstract number.
    /// </summary>
    /// <returns></returns>
    private static AbstractNum CreateNewAbstractNum() {
        AbstractNum newAbstractNum = new AbstractNum() { AbstractNumberId = nextAbstractNumberId };
        nextAbstractNumberId++;
        return newAbstractNum;
    }
}
