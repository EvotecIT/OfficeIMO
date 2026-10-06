namespace OfficeIMO.Workflows;

/// <summary>Bibliographic collection hierarchy levels from ONIX list 149.</summary>
public enum BookOnixCollectionLevel {
    /// <summary>Top-level collection (02).</summary>
    Collection,
    /// <summary>Subcollection (03).</summary>
    Subcollection,
    /// <summary>Subdivision of a subcollection (06).</summary>
    SubSubcollection
}

/// <summary>One collection title element. Supply Title, PartNumber, or both; list order determines display order.</summary>
public sealed record BookOnixCollectionTitleElement {
    /// <summary>Hierarchy level, independent of display order.</summary>
    public required BookOnixCollectionLevel Level { get; init; }
    /// <summary>Full title text, without a subtitle. Required when PartNumber is omitted.</summary>
    public string? Title { get; init; }
    /// <summary>Optional subtitle at this level.</summary>
    public string? Subtitle { get; init; }
    /// <summary>Optional part designation, including its caption, such as Volume II.</summary>
    public string? PartNumber { get; init; }
    /// <summary>Optional ONIX list 74 language of the title, subtitle and part designation.</summary>
    public string? LanguageCode { get; init; }
}

/// <summary>Publisher-declared frequency of successive products in a collection, from ONIX list 259.</summary>
public enum BookOnixCollectionFrequency {
    /// <summary>Explicitly unknown (u); distinct from omitting the assertion.</summary>
    Unknown,
    /// <summary>No fixed publication schedule (i).</summary>
    Irregular,
    /// <summary>Regularly scheduled, less frequently than once every two years (r).</summary>
    LessOftenThanBiennial,
    /// <summary>Once every two years (e).</summary>
    Biennial,
    /// <summary>Once per year (a).</summary>
    Annual,
    /// <summary>Twice per year (b).</summary>
    TwiceYearly,
    /// <summary>Three times per year (t).</summary>
    ThreeTimesYearly,
    /// <summary>Four times per year (q).</summary>
    Quarterly,
    /// <summary>Six times per year (s).</summary>
    EveryTwoMonths,
    /// <summary>Once per month (m).</summary>
    Monthly,
    /// <summary>Once every two weeks (f).</summary>
    Fortnightly,
    /// <summary>Once per week (w).</summary>
    Weekly,
    /// <summary>Daily or otherwise more frequently than weekly, at least twice per week (d).</summary>
    MoreOftenThanWeekly,
    /// <summary>No further publications are planned in the collection (x).</summary>
    NoFuturePublications
}
