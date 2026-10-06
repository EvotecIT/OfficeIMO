using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Optional font-program contract supplying static OpenType math design metrics.</summary>
public interface IOfficeMathFontProgram {
    /// <summary>Math constants, or null when absent, malformed, or not supported for the selected instance.</summary>
    OfficeMathFontConstants? MathConstants { get; }
}

/// <summary>
/// Immutable OpenType MATH constants. Device pixel corrections are deliberately omitted so
/// outline geometry scales consistently between screen and print. Variable math metrics are not qualified.
/// </summary>
public sealed class OfficeMathFontConstants {
    private readonly int[] _values;

    /// <summary>Creates a detached snapshot for a font provider. Supply every named constant in design units.</summary>
    public OfficeMathFontConstants(int unitsPerEm, IReadOnlyDictionary<OfficeMathConstant, int> values) {
        if (unitsPerEm <= 0) throw new ArgumentOutOfRangeException(nameof(unitsPerEm));
        if (values == null) throw new ArgumentNullException(nameof(values));
        if (values.Count != 56) throw new ArgumentException("Supply all 56 OpenType math constants.", nameof(values));
        _values = new int[56];
        for (int index = 0; index < _values.Length; index++) {
            if (!values.TryGetValue((OfficeMathConstant)index, out _values[index]))
                throw new ArgumentException("A named OpenType math constant is missing.", nameof(values));
            bool unsigned = index == 2 || index == 3;
            if (_values[index] < (unsigned ? 0 : short.MinValue) || _values[index] > (unsigned ? ushort.MaxValue : short.MaxValue))
                throw new ArgumentOutOfRangeException(nameof(values), "Math constants must fit their OpenType 16-bit representation.");
        }
        if (!ValidScriptScales(_values)) throw new ArgumentException("Script scales must be positive percentages at most 100, with the second no larger than the first.", nameof(values));
        UnitsPerEm = unitsPerEm;
    }

    internal OfficeMathFontConstants(int unitsPerEm, int[] values) {
        UnitsPerEm = unitsPerEm;
        _values = (int[])values.Clone();
    }

    internal static bool ValidScriptScales(int[] values) =>
        values[0] > 0 && values[0] <= 100 && values[1] > 0 && values[1] <= values[0];

    /// <summary>Design units per em used by the constants.</summary>
    public int UnitsPerEm { get; }

    /// <summary>
    /// Returns the unscaled design-unit value. ScriptPercentScaleDown,
    /// ScriptScriptPercentScaleDown and RadicalDegreeBottomRaisePercent are percentages.
    /// </summary>
    public int GetValue(OfficeMathConstant constant) {
        int index = (int)constant;
        if ((uint)index >= (uint)_values.Length) throw new ArgumentOutOfRangeException(nameof(constant));
        return _values[index];
    }
}
