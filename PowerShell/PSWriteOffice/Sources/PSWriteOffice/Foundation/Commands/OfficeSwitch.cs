using System;
using System.Collections.Generic;
using System.Management.Automation;

namespace PSWriteOffice.Foundation;

/// <summary>Preserves the distinction between omission and an explicit false switch.</summary>
public static class OfficeSwitch {
    /// <summary>Returns null for an unbound switch, otherwise its explicit boolean value.</summary>
    public static bool? Get(IDictionary<string, object> boundParameters, string name) {
        if (boundParameters == null) throw new ArgumentNullException(nameof(boundParameters));
        if (!boundParameters.TryGetValue(name, out object? value)) return null;
        return LanguagePrimitives.IsTrue(value);
    }
    /// <summary>Applies a NoFeature switch to a feature that is enabled by default; NoFeature:$false keeps it enabled.</summary>
    public static bool Enabled(IDictionary<string, object> boundParameters, string feature) => !(Get(boundParameters, "No" + feature) ?? false);
}
