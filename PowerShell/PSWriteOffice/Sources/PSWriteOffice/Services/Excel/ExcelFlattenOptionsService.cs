using System;
using System.Globalization;
using System.Management.Automation;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Excel;

/// <summary>Maps the shared PowerShell flattening policy to Excel's existing tabular input normalizer.</summary>
internal static class ExcelFlattenOptionsService {
    public static PowerShellObjectNormalizerOptions Create(OfficeObjectFlattenOptions? options, string? separator, CultureInfo? culture, Action<string, Exception>? reportError = null) {
        return new PowerShellObjectNormalizerOptions {
            CollectionSeparator = separator ?? ", ",
            DictionaryEntrySeparator = options?.DictionaryEntrySeparator ?? "; ",
            DictionaryKeyValueSeparator = options?.DictionaryKeyValueSeparator ?? ": ",
            MaxCollectionItems = options?.MaxCollectionItems ?? PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems,
            MaxNestingDepth = options?.MaxNestingDepth ?? PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth,
            IncludeUnexportableProperties = options?.IncludeUnreadableProperties ?? false,
            PropertyErrorAction = options?.PropertyConversionErrorAction ?? System.Management.Automation.ActionPreference.SilentlyContinue,
            Culture = culture ?? CultureInfo.InvariantCulture,
            PropertyErrorCallback = options?.PropertyConversionErrorAction == ActionPreference.Continue ? reportError : null,
            UnexportablePropertyValueFactory = static (_, exception) => $"Property export failed: {exception.Message}"
        };
    }
}
