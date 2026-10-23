using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeArabicTextShaper {
    /// <summary>Resolves joining features in logical source order without replacing Unicode characters.</summary>
    /// <remarks>Joining controls and transparent marks participate even when they have no painting glyph.</remarks>
    internal static string?[]? GetOpenTypeJoiningForms(string source) {
        if (!CanShapeAllJoiningCharacters(source)) return null;
        string logical = ToLogicalText(source);
        var result = new string?[logical.Length];
        for (int index = 0; index < logical.Length; index++) {
            if (!Forms.TryGetValue(logical[index], out ArabicForms forms)) continue;
            bool previous = forms.CanJoinPrevious && TryFindPrevious(logical, index, out ArabicForms before) && before.CanJoinNext;
            bool next = forms.CanJoinNext && TryFindNext(logical, index, out ArabicForms after) && after.CanJoinPrevious;
            result[index] = previous ? next ? "medi" : "fina" : next ? "init" : "isol";
        }
        return result;
    }
}
