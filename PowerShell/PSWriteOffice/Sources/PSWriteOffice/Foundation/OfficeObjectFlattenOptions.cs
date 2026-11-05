using System.Management.Automation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Foundation;

/// <summary>Controls dictionary and collection projection when exporting PowerShell objects to tables.</summary>
public sealed class OfficeObjectFlattenOptions {
    /// <summary>Separates entries when a dictionary is written into one cell.</summary>
    public string DictionaryEntrySeparator { get; set; } = "; ";

    /// <summary>Separates each dictionary key from its value.</summary>
    public string DictionaryKeyValueSeparator { get; set; } = ": ";

    /// <summary>Limits the number of items expanded or formatted from a collection.</summary>
    public int MaxCollectionItems { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems;

    /// <summary>Limits recursive traversal of nested dictionaries and objects.</summary>
    public int MaxNestingDepth { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth;

    /// <summary>Includes properties whose getters fail, using the normalizer's error-value policy.</summary>
    public bool IncludeUnreadableProperties { get; set; }

    /// <summary>Controls reporting of property getters or conversions that fail during projection.</summary>
    public ActionPreference PropertyConversionErrorAction { get; set; } = ActionPreference.SilentlyContinue;
}
