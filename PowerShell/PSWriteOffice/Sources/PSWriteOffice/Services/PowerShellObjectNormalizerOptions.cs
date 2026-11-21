using System.Globalization;
using System.Management.Automation;

namespace PSWriteOffice.Services;

/// <summary>Controls how PowerShell and CLR values are projected into Office table cells.</summary>
internal sealed class PowerShellObjectNormalizerOptions
{
    internal const int DefaultMaxCollectionItems = 1_048_575;
    internal const int DefaultMaxNestingDepth = 64;

    internal static readonly PowerShellObjectNormalizerOptions Default = new();

    public bool IncludeUnexportableProperties { get; set; }

    public ActionPreference PropertyErrorAction { get; set; } = ActionPreference.SilentlyContinue;

    public Action<string, Exception>? PropertyErrorCallback { get; set; }

    public Func<string, Exception, object?>? UnexportablePropertyValueFactory { get; set; }

    public bool NormalizeCollectionValues { get; set; } = true;

    public string CollectionSeparator { get; set; } = ", ";

    public string DictionaryEntrySeparator { get; set; } = "; ";

    public string DictionaryKeyValueSeparator { get; set; } = ": ";

    public int MaxCollectionItems { get; set; } = DefaultMaxCollectionItems;

    public int MaxNestingDepth { get; set; } = DefaultMaxNestingDepth;

    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    public bool FormatScalarValuesAsText { get; set; }

    /// <summary>
    /// Formats dates, times and floating-point numbers as text that reads back unchanged (ISO 8601 and round-trip
    /// numbers) when <see cref="FormatScalarValuesAsText"/> is set, as data formats such as CSV require.
    /// </summary>
    public bool RoundTripScalarText { get; set; }

    public static PowerShellObjectNormalizerOptions ForTable(
        string collectionSeparator,
        string dictionaryEntrySeparator,
        string dictionaryKeyValueSeparator,
        int maxCollectionItems = DefaultMaxCollectionItems,
        int maxNestingDepth = DefaultMaxNestingDepth)
    {
        return new PowerShellObjectNormalizerOptions
        {
            CollectionSeparator = collectionSeparator,
            DictionaryEntrySeparator = dictionaryEntrySeparator,
            DictionaryKeyValueSeparator = dictionaryKeyValueSeparator,
            MaxCollectionItems = maxCollectionItems,
            MaxNestingDepth = maxNestingDepth
        };
    }
}
