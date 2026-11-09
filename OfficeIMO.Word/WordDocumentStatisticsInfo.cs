namespace OfficeIMO.Word;

/// <summary>Detached snapshot of the native Word statistics counters.</summary>
public sealed class WordDocumentStatisticsInfo
{
    /// <summary>Creates a statistics snapshot from an OfficeIMO.Word statistics object.</summary>
    /// <param name="statistics">Document statistics read from OfficeIMO.Word.</param>
    internal WordDocumentStatisticsInfo(WordDocumentStatistics statistics)
    {
        statistics = statistics.CaptureForSnapshot();
        Pages = statistics.Pages;
        Paragraphs = statistics.Paragraphs;
        Words = statistics.Words;
        Images = statistics.Images;
        Tables = statistics.Tables;
        Charts = statistics.Charts;
        Shapes = statistics.Shapes;
        Bookmarks = statistics.Bookmarks;
        Lists = statistics.Lists;
        Characters = statistics.Characters;
        CharactersWithSpaces = statistics.CharactersWithSpaces;
    }

    /// <summary>Gets the estimated page count.</summary>
    public int Pages { get; }

    /// <summary>Gets the paragraph count.</summary>
    public int Paragraphs { get; }

    /// <summary>Gets the word count.</summary>
    public int Words { get; }

    /// <summary>Gets the image count.</summary>
    public int Images { get; }

    /// <summary>Gets the table count.</summary>
    public int Tables { get; }

    /// <summary>Gets the chart count.</summary>
    public int Charts { get; }

    /// <summary>Gets the shape count.</summary>
    public int Shapes { get; }

    /// <summary>Gets the bookmark count.</summary>
    public int Bookmarks { get; }

    /// <summary>Gets the list count.</summary>
    public int Lists { get; }

    /// <summary>Gets the character count without spaces.</summary>
    public int Characters { get; }

    /// <summary>Gets the character count including spaces.</summary>
    public int CharactersWithSpaces { get; }
}
